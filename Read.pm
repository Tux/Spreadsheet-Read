#!/pro/bin/perl

package Spreadsheet::Read;

=head1 NAME

 Spreadsheet::Read - Read the data from a spreadsheet

=head1 SYNOPSIS

 use Spreadsheet::Read;
 my $book  = ReadData ("test.csv", sep => ";");
 my $book  = ReadData ("test.sxc");
 my $book  = ReadData ("test.ods");
 my $book  = ReadData ("test.xls");
 my $book  = ReadData ("test.xlsx");
 my $book  = ReadData ($fh, parser => "xls");

 Spreadsheet::Read::add ($book, "sheet.csv");

 my $sheet = $book->[1];             # first datasheet
 my $cell  = $book->[1]{A3};         # content of field A3 of sheet 1
 my $cell  = $book->[1]{cell}[1][3]; # same, unformatted

 # OO API
 my $book = Spreadsheet::Read->new ("file.csv");
 my $sheet = $book->sheet (1);
 my $cell  = $sheet->cell ("A3");
 my $cell  = $sheet->cell (1, 3);

 $book->add ("test.xls");

=cut

use 5.8.1;
use strict;
use warnings;

our $VERSION = "0.80";
sub  Version { $VERSION }

use Carp;
use Exporter;
our @ISA       = qw( Exporter );
our @EXPORT    = qw( ReadData cell2cr cr2cell );
our @EXPORT_OK = qw( parses rows cellrow row add );

use Encode       qw( decode );
use File::Temp   qw( );
use Data::Dumper;

my @parsers = (
    [ csv  => "Text::CSV_XS",				"0.71"		],
    [ csv  => "Text::CSV_PP",				"1.17"		],
    [ csv  => "Text::CSV",				"1.17"		],
    [ ods  => "Spreadsheet::ReadSXC",			"0.20"		],
    [ sxc  => "Spreadsheet::ReadSXC",			"0.20"		],
    [ xls  => "Spreadsheet::ParseExcel",		"0.34"		],
    [ xlsx => "Spreadsheet::ParseXLSX",			"0.24"		],
    [ xlsx => "Spreadsheet::XLSX",			"0.13"		],
    [ prl  => "Spreadsheet::Perl",			""		],

    # Helper modules
    [ ios  => "IO::Scalar",				""		],
    [ dmp  => "Data::Peek",				""		],
    );
my %can = ( supports => { map { $_->[1] => $_->[2] } @parsers });
foreach my $p (@parsers) {
    my $format = $p->[0];
    $can{$format} and next;
    $can{$format} = "";
    my $preset = $ENV{"SPREADSHEET_READ_\U$format"} or next;
    my $min_version = $can{supports}{$preset};
    unless ($min_version) {
	# Catch weirdness like $SPREADSHEET_READ_XLSX = "DBD::Oracle"
	$can{$format} = "!$preset is not supported for the $format format";
	next;
	}
    if (eval "local \$_; require $preset" and not $@) {
	# forcing a parser should still check the version
	my $ok;
	my $has = $preset->VERSION;
	$has =~ s/_[0-9]+$//;			# Remove beta-part
	if ($min_version =~ m/^v([0-9.]+)/) {	# clumsy versions
	    my @min = split m/\./ => $1;
	    $has =~ s/^v//;
	    my @has = split m/\./ => $has;
	    $ok = (($has[0] * 1000 + $has[1]) * 1000 + $has[2]) >=
		  (($min[0] * 1000 + $min[1]) * 1000 + $min[2]);
	    }
	else {	# normal versions
	    $ok = $has >= $min_version;
	    }
	$ok or $preset = "!$preset";
	}
    else {
	$preset = "!$preset";
	}
    $can{$format} = $preset;
    }
delete $can{supports};
for (@parsers) {
    my ($flag, $mod, $vsn) = @$_;
    $can{$flag} and next;
    eval "require $mod; \$vsn and ${mod}->VERSION (\$vsn); \$can{\$flag} = '$mod'" or
	$_->[0] = "! Cannot use $mod version $vsn: $@";
    }
$can{sc} = __PACKAGE__;	# SquirelCalc is built-in

defined $Spreadsheet::ParseExcel::VERSION && $Spreadsheet::ParseExcel::VERSION < 0.61 and
    *Spreadsheet::ParseExcel::Workbook::get_active_sheet = sub { undef; };

my $debug = 0;
my %def_opts = (
    rc      => 1,
    cells   => 1,
    attr    => 0,
    clip    => undef, # $opt{cells};
    strip   => 0,
    pivot   => 0,
    dtfmt   => "yyyy-mm-dd", # Format 14
    debug   => 0,
    passwd  => undef,
    parser  => undef,
    sep     => undef,
    quote   => undef,
    label   => undef,
    );
my @def_attr = (
    type    => "text",
    fgcolor => undef,
    bgcolor => undef,
    font    => undef,
    size    => undef,
    format  => undef,
    halign  => "left",
    valign  => "top",
    bold    => 0,
    italic  => 0,
    uline   => 0,
    wrap    => 0,
    merged  => 0,
    hidden  => 0,
    locked  => 0,
    enc     => "utf-8", # $ENV{LC_ALL} // $ENV{LANG} // ...
    formula => undef,
    );

# Helper functions

sub _dump {
    my ($label, $ref) = @_;
    if ($can{dmp}) {
	print STDERR Data::Peek::DDumper ({ $label => $ref });
	}
    else {
	print STDERR Data::Dumper->Dump ([$ref], [$label]);
	}
    } # _dump

sub _parser {
    my $type = shift		or  return "";
    $type = lc $type;
    # Aliases and fullnames
    $type eq "excel"		and return "xls";
    $type eq "excel2007"	and return "xlsx";
    $type eq "oo"		and return "sxc";
    $type eq "ods"		and return "sxc";
    $type eq "openoffice"	and return "sxc";
    $type eq "libreoffice"	and return "sxc";
    $type eq "perl"		and return "prl";
    $type eq "squirelcalc"	and return "sc";
    return exists $can{$type} ? $type : "";
    } # _parser

sub new {
    my $class = shift;
    my $r = ReadData (@_) || [{
	parsers	=> [],
	error	=> undef,
	sheets	=> 0,
	sheet	=> { },
	}];
    bless $r => $class;
    } # new

# Spreadsheet::Read::parses ("csv") or die "Cannot parse CSV"
sub parses {
    ref $_[0] eq __PACKAGE__ and shift;
    my $type = _parser (shift)	or  return 0;
    if ($can{$type} =~ m/^!\s*(.*)/) {
	$@ = $1;
	return 0;
	}
    return $can{$type};
    } # parses

sub sheets {
    my $ctrl = shift->[0];
    my %s = %{$ctrl->{sheet}};
    wantarray ? sort { $s{$a} <=> $s{$b} } keys %s : $ctrl->{sheets};
    } # sheets

# col2label (4) => "D"
sub col2label {
    ref $_[0] eq __PACKAGE__ and shift;
    my $c = shift;
    defined $c && $c > 0 or return "";
    my $cell = "";
    while ($c) {
	use integer;

	substr $cell, 0, 0, chr (--$c % 26 + ord "A");
	$c /= 26;
	}
    $cell;
    } # col2label

# cr2cell (4, 18) => "D18"
# No prototype to allow 'cr2cell (@rowcol)'
sub cr2cell {
    ref $_[0] eq __PACKAGE__ and shift;
    my ($c, $r) = @_;
    defined $c && defined $r && $c > 0 && $r > 0 or return "";
    col2label ($c) . $r;
    } # cr2cell

# cell2cr ("D18") => (4, 18)
sub cell2cr {
    ref $_[0] eq __PACKAGE__ and shift;
    my ($cc, $r) = (uc ($_[0]||"") =~ m/^([A-Z]+)([0-9]+)$/) or return (0, 0);
    my $c = 0;
    while ($cc =~ s/^([A-Z])//) {
	$c = 26 * $c + 1 + ord ($1) - ord ("A");
	}
    ($c, $r);
    } # cell2cr

# my @row = cellrow ($book->[1], 1);
# my @row = $book->cellrow (1, 1);
sub cellrow {
    my $sheet = ref $_[0] eq __PACKAGE__ ? (shift)->[shift] : shift or return;
    ref     $sheet eq "HASH" && exists  $sheet->{cell}   or return;
    exists  $sheet->{maxcol} && exists  $sheet->{maxrow} or return;
    my $row   = shift or return;
    $row > 0 && $row <= $sheet->{maxrow} or return;
    my $s = $sheet->{cell};
    map { $s->[$_][$row] } 1..$sheet->{maxcol};
    } # cellrow

# my @row = row ($book->[1], 1);
sub row {
    my $sheet = ref $_[0] eq __PACKAGE__ ? (shift)->[shift] : shift or return;
    ref     $sheet eq "HASH" && exists  $sheet->{cell}   or return;
    exists  $sheet->{maxcol} && exists  $sheet->{maxrow} or return;
    my $row   = shift or return;
    $row > 0 && $row <= $sheet->{maxrow} or return;
    map { $sheet->{cr2cell ($_, $row)} } 1..$sheet->{maxcol};
    } # row

# Convert {cell}'s [column][row] to a [row][column] list
# my @rows = rows ($book->[1]);
sub rows {
    my $sheet = ref $_[0] eq __PACKAGE__ ? (shift)->[shift] : shift or return;
    ref    $sheet eq "HASH" && exists $sheet->{cell}   or return;
    exists $sheet->{maxcol} && exists $sheet->{maxrow} or return;
    my $s = $sheet->{cell};

    map {
	my $r = $_;
	[ map { $s->[$_][$r] } 1..$sheet->{maxcol} ];
	} 1..$sheet->{maxrow};
    } # rows

sub sheet {
    my ($book, $sheet) = @_;
    $book && $sheet or return;
    my $class = "Spreadsheet::Read::Sheet";
    $sheet =~ m/^[0-9]+$/ && $sheet >= 1 && $sheet <= $book->[0]{sheets} and
	return bless $book->[$sheet]			=> $class;
    exists $book->[0]{sheet}{$sheet} and
	return bless $book->[$book->[0]{sheet}{$sheet}]	=> $class;
    foreach my $idx (1 .. $book->[0]{sheets}) {
	$book->[$idx]{label} eq $sheet and
	    return bless $book->[$idx]			=> $class;
	}
    return;
    } # sheet

# If option "clip" is set, remove the trailing rows and
# columns in each sheet that contain no visible data
sub _clipsheets {
    my ($opt, $ref) = @_;

    $ref->[0]{sheets} or return $ref;

    my ($rc, $cl)      = ($opt->{rc},   $opt->{cells});
    my ($oc, $os, $oa) = ($opt->{clip}, $opt->{strip}, $opt->{attr});

    # Strip leading/trailing spaces
    if ($os || $oc) {
	foreach my $sheet (1 .. $ref->[0]{sheets}) {
	    $ref->[$sheet]{indx} = $sheet;
	    my $ss = $ref->[$sheet];
	    $ss->{maxrow} && $ss->{maxcol} or next;
	    my ($mc, $mr) = (0, 0);
	    foreach my $row (1 .. $ss->{maxrow}) {
		foreach my $col (1 .. $ss->{maxcol}) {
		    if ($rc) {
			defined $ss->{cell}[$col][$row] or next;
			$os & 2 and $ss->{cell}[$col][$row] =~ s/\s+$//;
			$os & 1 and $ss->{cell}[$col][$row] =~ s/^\s+//;
			if (length $ss->{cell}[$col][$row]) {
			    $col > $mc and $mc = $col;
			    $row > $mr and $mr = $row;
			    }
			}
		    if ($cl) {
			my $cell = cr2cell ($col, $row);
			defined $ss->{$cell} or next;
			$os & 2 and $ss->{$cell} =~ s/\s+$//;
			$os & 1 and $ss->{$cell} =~ s/^\s+//;
			if (length $ss->{$cell}) {
			    $col > $mc and $mc = $col;
			    $row > $mr and $mr = $row;
			    }
			}
		    }
		}

	    $oc && ($mc < $ss->{maxcol} || $mr < $ss->{maxrow}) or next;

	    # Remove trailing empty columns
	    foreach my $col (($mc + 1) .. $ss->{maxcol}) {
		$rc and undef $ss->{cell}[$col];
		$oa and undef $ss->{attr}[$col];
		$cl or next;
		my $c = col2label ($col);
		delete $ss->{"$c$_"} for 1 .. $ss->{maxrow};
		}

	    # Remove trailing empty rows
	    foreach my $row (($mr + 1) .. $ss->{maxrow}) {
		foreach my $col (1 .. $mc) {
		    $cl and delete $ss->{cr2cell ($col, $row)};
		    $rc and undef  $ss->{cell}   [$col][$row];
		    $oa and undef  $ss->{attr}   [$col][$row];
		    }
		}

	    ($ss->{maxrow}, $ss->{maxcol}) = ($mr, $mc);
	    }
	}

    if ($opt->{pivot}) {
	foreach my $sheet (1 .. $ref->[0]{sheets}) {
	    my $ss = $ref->[$sheet];
	    $ss->{maxrow} || $ss->{maxcol} or next;
	    my $mx = $ss->{maxrow} > $ss->{maxcol} ? $ss->{maxrow} : $ss->{maxcol};
	    foreach my $row (2 .. $mx) {
		foreach my $col (1 .. ($row - 1)) {
		    $opt->{rc}    and
			($ss->{cell}[$col][$row], $ss->{cell}[$row][$col]) =
			($ss->{cell}[$row][$col], $ss->{cell}[$col][$row]);
		    $opt->{cells} and
		        ($ss->{cr2cell ($col, $row)}, $ss->{cr2cell ($row, $col)}) =
		        ($ss->{cr2cell ($row, $col)}, $ss->{cr2cell ($col, $row)});
		    }
		}
	    ($ss->{maxcol}, $ss->{maxrow}) = ($ss->{maxrow}, $ss->{maxcol});
	    }
	}

    $ref;
    } # _clipsheets

# Convert a single color (index) to a color
sub _xls_color {
    my $clr = shift;
    defined $clr		or  return undef;
    $clr eq "#000000"		and return undef;
    $clr =~ m/^#[0-9a-fA-F]+$/	and return lc $clr;
    $clr == 0 || $clr == 32767	and return undef; # Default fg color
    return "#" . lc Spreadsheet::ParseExcel->ColorIdxToRGB ($clr);
    } # _xls_color

# Convert a fill [ $pattern, $front_color, $back_color ] to a single background
sub _xls_fill {
    my ($p, $fg, $bg) = @_;
    defined $p			or  return undef;
    $p == 32767			and return undef; # Default fg color
    $p == 0 && !defined $bg	and return undef; # No fill bg color
    $p == 1			and return _xls_color ($fg);
    $bg < 8 || $bg > 63		and return undef; # see Workbook.pm#106
    return _xls_color ($bg);
    } # _xls_fill

sub ReadData {
    my $txt = shift	or  return;

    my %opt;
    if (@_) {
	   if (ref $_[0] eq "HASH")  { %opt = %{shift @_} }
	elsif (@_ % 2 == 0)          { %opt = @_          }
	}

    exists $opt{rc}	or $opt{rc}	= $def_opts{rc};
    exists $opt{cells}	or $opt{cells}	= $def_opts{cells};
    exists $opt{attr}	or $opt{attr}	= $def_opts{attr};
    exists $opt{clip}	or $opt{clip}	= $opt{cells};
    exists $opt{strip}	or $opt{strip}	= $def_opts{strip};
    exists $opt{dtfmt}	or $opt{dtfmt}	= $def_opts{dtfmt};

    # $debug = $opt{debug} // 0;
    $debug = defined $opt{debug} ? $opt{debug} : $def_opts{debug};
    $debug > 4 and _dump (Options => \%opt);

    my %parser_opts = map { $_ => $opt{$_} }
		      grep { !exists $def_opts{$_} }
		      keys %opt;

    my $_parser = _parser ($opt{parser});

    my $io_ref = ref ($txt) =~ m/GLOB|IO/ ? $txt : undef;
    my $io_fil = $io_ref ? 0 : $txt =~ m/\0/ ? 0
			     : do { no warnings "newline"; -f $txt };
    my $io_txt = $io_ref || $io_fil ? 0 : 1;

    $io_fil && ! -s $txt  and return;
    $io_ref && eof ($txt) and return;

    if ($opt{parser} ? $_parser eq "csv" : ($io_fil && $txt =~ m/\.(csv)$/i)) {
	$can{csv} or croak "CSV parser not installed";

	my $label = defined $opt{label} ? $opt{label} : $io_fil ? $txt : "IO";

	$debug and print STDERR "Opening CSV $label using $can{csv}-", $can{csv}->VERSION, "\n";

	my @data = (
	    {	type	=> "csv",
		parser	=> $can{csv},
		version	=> $can{csv}->VERSION,
		parsers	=> [ {
		    type	=> "csv",
		    parser	=> $can{csv},
		    version	=> $can{csv}->VERSION,
		    }],
		error	=> undef,
		quote	=> '"',
		sepchar	=> ',',
		sheets	=> 1,
		sheet	=> { $label => 1 },
		},
	    {	parser	=> 0,
		label	=> $label,
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		merged	=> [],
		active  => 1,
		},
	    );

	my ($sep, $quo, $in) = (",", '"');
	defined $opt{sep}   and $sep = $opt{sep};
	defined $opt{quote} and $quo = $opt{quote};
	$debug > 8 and _dump (debug => {
	    data => \@data, txt => $txt, io_ref => $io_ref, io_fil => $io_fil });
	if ($io_fil) {
	    unless (defined $opt{quote} && defined $opt{sep}) {
		open $in, "<", $txt or return;
		my $l1 = <$in>;

		$quo = defined $opt{quote} ? $opt{quote} : '"';
		$sep = # If explicitly set, use it
		   defined $opt{sep} ? $opt{sep} :
		       # otherwise start auto-detect with quoted strings
		       $l1 =~ m/["0-9];["0-9;]/  ? ";"  :
		       $l1 =~ m/["0-9],["0-9,]/  ? ","  :
		       $l1 =~ m/["0-9]\t["0-9,]/ ? "\t" :
		       # If neither, then for unquoted strings
		       $l1 =~ m/\w;[\w;]/        ? ";"  :
		       $l1 =~ m/\w,[\w,]/        ? ","  :
		       $l1 =~ m/\w\t[\w,]/       ? "\t" :
						   ","  ;
		close $in;
		}
	    open $in, "<", $txt or return;
	    }
	elsif ($io_ref) {
	    $in = $txt;
	    }
	elsif (ref $txt eq "SCALAR") {
	    open $in, "<", $txt  or croak "Cannot open input: $!";
	    }
	elsif ($txt =~ m/[\r\n,;]/) {
	    open $in, "<", \$txt or croak "Cannot open input: $!";
	    }
	else {
	    warn "Input type ", ref $txt,
		" might not be supported. Please file a ticket\n";
	    $in = $txt;	# Now pray ...
	    }
	$debug > 1 and print STDERR "CSV sep_char '$sep', quote_char '$quo'\n";
	my $csv = $can{csv}->new ({
	    %parser_opts,

	    sep_char       => ($data[0]{sepchar} = $sep),
	    quote_char     => ($data[0]{quote}   = $quo),
	    keep_meta_info => 1,
	    binary         => 1,
	    auto_diag      => 1,
	    }) or croak "Cannot create a csv ('$sep', '$quo') parser!";

	while (my $row = $csv->getline ($in)) {
	    my @row = @$row or last;

	    my $r = ++$data[1]{maxrow};
	    @row > $data[1]{maxcol} and $data[1]{maxcol} = @row;
	    foreach my $c (0 .. $#row) {
		my $val = $row[$c];
		my $cell = cr2cell ($c + 1, $r);
		$opt{rc}    and $data[1]{cell}[$c + 1][$r] = $val;
		$opt{cells} and $data[1]{$cell} = $val;
		$opt{attr}  and $data[1]{attr}[$c + 1][$r] = { @def_attr };
		}
	    }
	$csv->eof () or $data[0]{error} = [ $csv->error_diag ];
	close $in;

	for (@{$data[1]{cell}}) {
	    defined or $_ = [];
	    }
	return _clipsheets \%opt, [ @data ];
	}

    if ($io_txt) { # && $_parser !~ m/^xlsx?$/) {
	if (    # /etc/magic: Microsoft Office Document
		$txt =~ m{\A(\376\067\0\043
			    |\320\317\021\340\241\261\032\341
			    |\333\245-\0\0\0)}x
		# /usr/share/misc/magic
	     || $txt =~ m{\A.{2080}Microsoft Excel 5.0 Worksheet}
	     || $txt =~ m{\A\x09\x04\x06\x00\x00\x00\x10\x00}
		) {
	    $can{xls} or croak "Spreadsheet::ParseExcel not installed";
	    my $tmpfile;
	    if ($can{ios}) { # Do not use a temp file if IO::Scalar is available
		$tmpfile = \$txt;
		}
	    else {
		$tmpfile = File::Temp->new (SUFFIX => ".xls", UNLINK => 1);
		binmode $tmpfile;
		print   $tmpfile $txt;
		close   $tmpfile;
		}
	    open $io_ref, "<", $tmpfile or return;
	    $io_txt = 0;
	    $_parser = _parser ($opt{parser} = "xls");
	    }
	elsif ( # /usr/share/misc/magic
		$txt =~ m{\APK\003\004.{4,30}(?:\[Content_Types\]\.xml|_rels/\.rels)}
		) {
	    $can{xlsx} or croak "XLSX parser not installed";
	    my $tmpfile;
	    if ($can{ios}) { # Do not use a temp file if IO::Scalar is available
		$tmpfile = \$txt;
		}
	    else {
		$tmpfile = File::Temp->new (SUFFIX => ".xlsx", UNLINK => 1);
		binmode $tmpfile;
		print   $tmpfile $txt;
		close   $tmpfile;
		}
	    open $io_ref, "<", $tmpfile or return;
	    $io_txt = 0;
	    $_parser = _parser ($opt{parser} = "xlsx");
	    }
	}
    if ($opt{parser} ? $_parser =~ m/^xlsx?$/
		     : ($io_fil && $txt =~ m/\.(xlsx?)$/i && ($_parser = $1))) {
	my $parse_type = $_parser =~ m/x$/i ? "XLSX" : "XLS";
	my $parser = $can{lc $parse_type} or
	    croak "Parser for $parse_type is not installed";
	$debug and print STDERR "Opening $parse_type ", $io_ref ? "<REF>" : $txt,
	    " using $parser-", $can{lc $parse_type}->VERSION, "\n";
	$opt{passwd} and $parser_opts{Password} = $opt{passwd};
	my $oBook = eval {
	    $io_ref
	      ? $parse_type eq "XLSX"
		? $can{xlsx} =~ m/::XLSX$/
		? $parser->new ($io_ref)
		: $parser->new (%parser_opts)->parse ($io_ref)
		: $parser->new (%parser_opts)->Parse ($io_ref)
	      : $parse_type eq "XLSX"
		? $can{xlsx} =~ m/::XLSX$/
		? $parser->new ($txt)
		: $parser->new (%parser_opts)->parse ($txt)
		: $parser->new (%parser_opts)->Parse ($txt);
	    };
	unless ($oBook) {
	    # cleanup will fail on folders with spaces.
	    (my $msg = $@) =~ s/ at \S+ line \d+.*//s;
	    croak "$parse_type parser cannot parse data: $msg";
	    }
	$debug > 8 and _dump (oBook => $oBook);

	# WorkBook keys:
	# aColor         _CurSheet      Format         SheetCount
	# ActiveSheet    _CurSheet_     FormatStr      _skip_chart
	# Author         File           NotSetCell     _string_contin
	# BIFFVersion    Flg1904        Object         Version
	# _buffer        FmtClass       PkgStr         Worksheet
	# CellHandler    Font           _previous_info

	my @data = ( {
	    type	=> lc $parse_type,
	    parser	=> $can{lc $parse_type},
	    version	=> $can{lc $parse_type}->VERSION,
	    parsers	=> [{
		type	=> lc $parse_type,
		parser	=> $can{lc $parse_type},
		version	=> $can{lc $parse_type}->VERSION,
		}],
	    error	=> undef,
	    sheets	=> $oBook->{SheetCount} || 0,
	    sheet	=> {},
	    } );
	# $debug and $data[0]{_parser} = $oBook;
	# Overrule the default date format strings
	my %def_fmt = (
	    0x0E	=> lc $opt{dtfmt},	# m-d-yy
	    0x0F	=> "d-mmm-yyyy",	# d-mmm-yy
	    0x11	=> "mmm-yyyy",		# mmm-yy
	    0x16	=> "yyyy-mm-dd hh:mm",	# m-d-yy h:mm
	    );
	$oBook->{FormatStr}{$_} = $def_fmt{$_} for keys %def_fmt;
	my $oFmt = $parse_type eq "XLSX"
	    ? $can{xlsx} =~ m/::XLSX$/
		? Spreadsheet::XLSX::Fmt2007->new
		: Spreadsheet::ParseExcel::FmtDefault->new
	    :     Spreadsheet::ParseExcel::FmtDefault->new;

	$debug and print STDERR "\t$data[0]{sheets} sheets\n";
	my $active_sheet = $oBook->get_active_sheet
			|| $oBook->{ActiveSheet}
			|| $oBook->{SelectedSheet};
	my $current_sheet = 0;
	foreach my $oWkS (@{$oBook->{Worksheet}}) {
	    $current_sheet++;
	    $opt{clip} and !defined $oWkS->{Cells} and next; # Skip empty sheets
	    my %sheet = (
		parser	=> 0,
		label	=> $oWkS->{Name},
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		merged  => [],
		active	=> 0,
		);
	    # $debug and $sheet{_parser} = $oWkS;
	    defined $sheet{label}  or  $sheet{label}  = "-- unlabeled --";
	    exists $oWkS->{MinRow} and $sheet{minrow} = $oWkS->{MinRow} + 1;
	    exists $oWkS->{MaxRow} and $sheet{maxrow} = $oWkS->{MaxRow} + 1;
	    exists $oWkS->{MinCol} and $sheet{mincol} = $oWkS->{MinCol} + 1;
	    exists $oWkS->{MaxCol} and $sheet{maxcol} = $oWkS->{MaxCol} + 1;
	    $sheet{merged} = [
		map  {  $_->[0] }
		sort {  $a->[1] cmp $b->[1] }
		map  {[ $_, pack "NNNN", @$_          ]}
		map  {[ map { $_ + 1 } @{$_}[1,0,3,2] ]}
		@{$oWkS->get_merged_areas || []}];
	    my $sheet_idx = 1 + @data;
	    $debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} x $sheet{maxcol}\n";
	    if (defined $active_sheet) {
		# _SheetNo is 0-based
		my $sheet_no = defined $oWkS->{_SheetNo} ? $oWkS->{_SheetNo} : $current_sheet - 1;
		$sheet_no eq $active_sheet and $sheet{active} = 1;
		}
	    # Sheet keys:
	    # _Book          FooterMargin   MinCol         RightMargin
	    # BottomMargin   FooterMergin   MinRow         RightMergin
	    # BottomMergin   HCenter        Name           RowHeight
	    # Cells          Header         NoColor        RowHidden
	    # ColFmtNo       HeaderMargin   NoOrient       Scale
	    # ColHidden      HeaderMergin   NoPls          SheetHidden
	    # ColWidth       Kind           Notes          _SheetNo
	    # Copis          Landscape      PageFit        SheetType
	    # DefColWidth    LeftMargin     PageStart      SheetVersion
	    # DefRowHeight   LeftMergin     PaperSize      TopMargin
	    # Draft          LeftToRight    _Pos           TopMergin
	    # FitHeight      MaxCol         PrintGrid      UsePage
	    # FitWidth       MaxRow         PrintHeaders   VCenter
	    # Footer         MergedArea     Res            VRes
	    if (exists $oWkS->{MinRow}) {
		my $hiddenRows = $oWkS->{RowHidden} || [];
		my $hiddenCols = $oWkS->{ColHidden} || [];
		if ($opt{clip}) {
		    my ($mr, $mc) = (-1, -1);
		    foreach my $r ($oWkS->{MinRow} .. $sheet{maxrow}) {
			foreach my $c ($oWkS->{MinCol} .. $sheet{maxcol}) {
			    my $oWkC = $oWkS->{Cells}[$r][$c] or next;
			    defined (my $val = $oWkC->{Val})  or next;
			    $val eq "" and next;
			    $r > $mr and $mr = $r;
			    $c > $mc and $mc = $c;
			    }
			}
		    ($sheet{maxrow}, $sheet{maxcol}) = ($mr + 1, $mc + 1);
		    }
		foreach my $r ($oWkS->{MinRow} .. $sheet{maxrow}) {
		    foreach my $c ($oWkS->{MinCol} .. $sheet{maxcol}) {
			my $oWkC = $oWkS->{Cells}[$r][$c] or next;
			#defined (my $val = $oWkC->{Val}) or next;
			my $val = $oWkC->{Val};
			if (defined $val and my $enc = $oWkC->{Code}) {
			    $enc eq "ucs2" and $val = decode ("utf-16be", $val);
			    }
			my $cell = cr2cell ($c + 1, $r + 1);
			$opt{rc} and $sheet{cell}[$c + 1][$r + 1] = $val;	# Original

			my $fmt;
			my $FmT = $oWkC->{Format};
			if ($FmT) {
			    unless (ref $FmT) {
				$fmt = $FmT;
				$FmT = {};
				}
			    }
			else {
			    $FmT = {};
			    }
			foreach my $attr (qw( AlignH AlignV FmtIdx Hidden Lock
					      Wrap )) {
			    exists $FmT->{$attr} or $FmT->{$attr} = 0;
			    }
			exists $FmT->{Fill} or $FmT->{Fill} = [ 0 ];
			exists $FmT->{Font} or $FmT->{Font} = undef;

			unless (defined $fmt) {
			    $fmt = $FmT->{FmtIdx}
			       ? $oBook->{FormatStr}{$FmT->{FmtIdx}}
			       : undef;
			    }
			if ($oWkC->{Type} eq "Numeric") {
			    # Fixed in 0.33 and up
			    # see Spreadsheet/ParseExcel/FmtDefault.pm
			    $FmT->{FmtIdx} == 0x0e ||
			    $FmT->{FmtIdx} == 0x0f ||
			    $FmT->{FmtIdx} == 0x10 ||
			    $FmT->{FmtIdx} == 0x11 ||
			    $FmT->{FmtIdx} == 0x16 ||
			    (defined $fmt && $fmt =~ m{^[dmy][-\\/dmy]*$}) and
				$oWkC->{Type} = "Date";
			    $FmT->{FmtIdx} == 0x09 ||
			    $FmT->{FmtIdx} == 0x0a ||
			    (defined $fmt && $fmt =~ m{^0+\.0+%$}) and
				$oWkC->{Type} = "Percentage";
			    }
			defined $fmt and $fmt =~ s/\\//g;
			$opt{cells} and	# Formatted value
			    $sheet{$cell} = defined $val ? $FmT && exists $def_fmt{$FmT->{FmtIdx}}
				? $oFmt->ValFmt ($oWkC, $oBook)
				: $oWkC->Value : undef;
			if ($opt{attr}) {
			    my $FnT = $FmT->{Font};
			    my $fmi = $FmT->{FmtIdx}
			       ? $oBook->{FormatStr}{$FmT->{FmtIdx}}
			       : undef;
			    $fmi and $fmi =~ s/\\//g;
			    $sheet{attr}[$c + 1][$r + 1] = {
				@def_attr,

				type    => lc $oWkC->{Type},
				enc     => $oWkC->{Code},
				merged  => (defined $oWkC->{Merged} ? $oWkC->{Merged} : $oWkC->is_merged) || 0,
				hidden  => ($hiddenRows->[$r] || $hiddenCols->[$c] ? 1 :
					    defined $oWkC->{Hidden} ? $oWkC->{Hidden} : $FmT->{Hidden})   || 0,
				locked  => $FmT->{Lock}     || 0,
				format  => $fmi,
				halign  => [ undef, qw( left center right
					   fill justify ), undef,
					   "equal_space" ]->[$FmT->{AlignH}],
				valign  => [ qw( top center bottom justify
					   equal_space )]->[$FmT->{AlignV}],
				wrap    => $FmT->{Wrap},
				font    => $FnT->{Name},
				size    => $FnT->{Height},
				bold    => $FnT->{Bold},
				italic  => $FnT->{Italic},
				uline   => $FnT->{Underline},
				fgcolor => _xls_color ($FnT->{Color}),
				bgcolor => _xls_fill  (@{$FmT->{Fill}}),
				formula => $oWkC->{Formula},
				};
			    #_dump "cell", $sheet{attr}[$c + 1][$r + 1];
			    }
			}
		    }
		}
	    for (@{$sheet{cell}}) {
		defined or $_ = [];
		}
	    push @data, { %sheet };
#	    $data[0]{sheets}++;
	    if ($sheet{label} eq "-- unlabeled --") {
		$sheet{label} = "";
		}
	    else {
		$data[0]{sheet}{$sheet{label}} = $#data;
		}
	    }
	return _clipsheets \%opt, [ @data ];
	}

    if ($opt{parser} ? _parser ($opt{parser}) eq "sc"
		     : $io_fil
			 ? $txt =~ m/\.sc$/
			 : $txt =~ m/^# .*SquirrelCalc/) {
	if ($io_ref) {
	    local $/;
	    my $x = <$txt>;
	    $txt = $x;
	    }
	elsif ($io_fil) {
	    local $/;
	    open my $sc, "<", $txt or return;
	    $txt = <$sc>;
	    close   $sc;
	    }
	$txt =~ m/\S/ or return;
	my $label = defined $opt{label} ? $opt{label} : "sheet";
	my @data = (
	    {	type	=> "sc",
		parser	=> "Spreadsheet::Read",
		version	=> $VERSION,
		parsers	=> [{
		    type	=> "sc",
		    parser	=> "Spreadsheet::Read",
		    version	=> $VERSION,
		    }],
		error	=> undef,
		sheets	=> 1,
		sheet	=> { $label => 1 },
		},
	    {	parser	=> 0,
		label	=> $label,
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		merged  => [],
		active  => 1,
		},
	    );

	for (split m/\s*[\r\n]\s*/, $txt) {
	    if (m/^dimension.*of ([0-9]+) rows.*of ([0-9]+) columns/i) {
		@{$data[1]}{qw(maxrow maxcol)} = ($1, $2);
		next;
		}
	    s/^r([0-9]+)c([0-9]+)\s*=\s*// or next;
	    my ($c, $r) = map { $_ + 1 } $2, $1;
	    if (m/.* \{(.*)}$/ or m/"(.*)"/) {
		my $cell = cr2cell ($c, $r);
		$opt{rc}    and $data[1]{cell}[$c][$r] = $1;
		$opt{cells} and $data[1]{$cell} = $1;
		$opt{attr}  and $data[1]{attr}[$c + 1][$r] = { @def_attr };
		next;
		}
	    # Now only formula's remain. Ignore for now
	    # r67c7 = [P2L] 2*(1000*r67c5-60)
	    }
	for (@{$data[1]{cell}}) {
	    defined or $_ = [];
	    }
	return _clipsheets \%opt, [ @data ];
	}

    if ($opt{parser} ? _parser ($opt{parser}) eq "sxc"
		     : ($txt =~ m/^<\?xml/ or -f $txt)) {
	$can{sxc} or croak "Spreadsheet::ReadSXC not installed";
	ref $txt and
	    croak ("Sorry, references as input are not (yet) supported by Spreadsheet::ReadSXC");

	my $using = "using $can{sxc}-" . $can{sxc}->VERSION;
	my $sxc_options = { %parser_opts, OrderBySheet => 1 }; # New interface 0.20 and up
	my $sxc;
	   if ($txt =~ m/\.(sxc|ods)$/i) {
	    $debug and print STDERR "Opening \U$1\E $txt $using\n";
	    $sxc = Spreadsheet::ReadSXC::read_sxc      ($txt, $sxc_options)	or  return;
	    }
	elsif ($txt =~ m/\.xml$/i) {
	    $debug and print STDERR "Opening XML $txt $using\n";
	    $sxc = Spreadsheet::ReadSXC::read_xml_file ($txt, $sxc_options)	or  return;
	    }
	# need to test on pattern to prevent stat warning
	# on filename with newline
	elsif ($txt !~ m/^<\?xml/i and -f $txt) {
	    $debug and print STDERR "Opening XML $txt $using\n";
	    open my $f, "<", $txt	or  return;
	    local $/;
	    $txt = <$f>;
	    close $f;
	    }
	!$sxc && $txt =~ m/^<\?xml/i and
	    $sxc = Spreadsheet::ReadSXC::read_xml_string ($txt, $sxc_options);
	$debug > 8 and _dump (sxc => $sxc);
	if ($sxc) {
	    my @data = ( {
		type	=> "sxc",
		parser	=> "Spreadsheet::ReadSXC",
		version	=> $Spreadsheet::ReadSXC::VERSION,
		parsers	=> [{
		    type	=> "sxc",
		    parser	=> "Spreadsheet::ReadSXC",
		    version	=> $Spreadsheet::ReadSXC::VERSION,
		    }],
		error	=> undef,
		sheets	=> 0,
		sheet	=> {},
		} );
	    my @sheets = ref $sxc eq "HASH"	# < 0.20
		? map {
		    {   label => $_,
			data  => $sxc->{$_},
			}
		    } keys %$sxc
		: @{$sxc};
	    foreach my $sheet (@sheets) {
		my @sheet = @{$sheet->{data}};
		my %sheet = (
		    parser => 0,
		    label  => $sheet->{label},
		    maxrow => scalar @sheet,
		    maxcol => 0,
		    cell   => [],
		    attr   => [],
		    merged => [],
		    active => 0,
		    );
		my $sheet_idx = 1 + @data;
		$debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} rows\n";
		foreach my $r (0 .. $#sheet) {
		    my @row = @{$sheet[$r]} or next;
		    foreach my $c (0 .. $#row) {
			defined (my $val = $row[$c]) or next;
			my $C = $c + 1;
			$C > $sheet{maxcol} and $sheet{maxcol} = $C;
			my $cell = cr2cell ($C, $r + 1);
			$opt{rc}    and $sheet{cell}[$C][$r + 1] = $val;
			$opt{cells} and $sheet{$cell} = $val;
			$opt{attr}  and $sheet{attr}[$C][$r + 1] = { @def_attr };
			}
		    }
		for (@{$sheet{cell}}) {
		    defined or $_ = [];
		    }
		$debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} x $sheet{maxcol}\n";
		push @data, { %sheet };
		$data[0]{sheets}++;
		$data[0]{sheet}{$sheet->{label}} = $#data;
		}
	    return _clipsheets \%opt, [ @data ];
	    }
	}

    return;
    } # ReadData

sub add {
    my $book = shift;
    my $r = ReadData (@_) or return;
    $book && (ref $book eq "ARRAY" ||
	      ref $book eq __PACKAGE__) && $book->[0]{sheets} or return $r;

    my $c1 = $book->[0];
    my $c2 = $r->[0];

    unless ($c1->{parsers}) {
	$c1->{parsers}[0]{$_} = $c1->{$_} for qw( type parser version );
	$book->[$_]{parser} = 0 for 1 .. $c1->{sheets};
	}
    my ($pidx) = (grep { my $p = $c1->{parsers}[$_];
	$p->{type}    eq $c2->{type}   &&
	$p->{parser}  eq $c2->{parser} &&
	$p->{version} eq $c2->{version} } 0 .. $#{$c1->{parsers}});
    unless (defined $pidx) {
	$pidx = scalar @{$c1->{parsers}};
	$c1->{parsers}[$pidx]{$_} = $c2->{$_} for qw( type parser version );
	}

    foreach my $sn (sort { $c2->{sheet}{$a} <=> $c2->{sheet}{$b} } keys %{$c2->{sheet}}) {
	my $s = $sn;
	my $v = 2;
	while (exists $c1->{sheet}{$s}) {
	    $s = $sn."[".$v++."]";
	    }
	$c1->{sheet}{$s} = $c1->{sheets} + $c2->{sheet}{$sn};
	$r->[$c2->{sheet}{$sn}]{parser} = $pidx;
	push @$book, $r->[$c2->{sheet}{$sn}];
	}
    $c1->{sheets} += $c2->{sheets};

    return $book;
    } # add

package Spreadsheet::Read::Attribute;

use Carp;
use vars qw( $AUTOLOAD );

sub AUTOLOAD {
    my $self = shift;
    (my $attr = $AUTOLOAD) =~ s/.*:://;
    $self->{$attr};
    } # AUTOLOAD

package Spreadsheet::Read::Sheet;

sub cell {
    my ($sheet, @id) = @_;
    @id == 2 && $id[0] =~ m/^[0-9]+$/ && $id[1] =~ m/^[0-9]+$/ and
	return $sheet->{cell}[$id[0]][$id[1]];
    @id && $id[0] && exists $sheet->{$id[0]} and
	return $sheet->{$id[0]};
    } # cell

sub attr {
    my ($sheet, @id) = @_;
    my $class = "Spreadsheet::Read::Attribute";
    @id == 2 && $id[0] =~ m/^[0-9]+$/ && $id[1] =~ m/^[0-9]+$/ and
	return bless $sheet->{attr}[$id[0]][$id[1]] => $class;
    if (@id && $id[0] && exists $sheet->{$id[0]}) {
	my ($c, $r) = $sheet->cell2cr ($id[0]);
	return bless $sheet->{attr}[$c][$r] => $class;
	}
    undef;
    } # attr

sub maxrow {
    my $sheet = shift;
    return $sheet->{maxrow};
    } # maxrow

sub maxcol {
    my $sheet = shift;
    return $sheet->{maxcol};
    } # maxrow

sub col2label {
    $_[0] =~ m/::/ and shift; # class unused
    return Spreadsheet::Read::col2label (@_);
    } # col2label

sub cr2cell {
    $_[0] =~ m/::/ and shift; # class unused
    return Spreadsheet::Read::cr2cell (@_);
    } # cr2cell

sub cell2cr {
    $_[0] =~ m/::/ and shift; # class unused
    return Spreadsheet::Read::cell2cr (@_);
    } # cell2cr

sub label {
    my ($sheet, $label) = @_;
    defined $label and $sheet->{label} = $label;
    return $sheet->{label};
    } # label

sub active {
    my $sheet = shift;
    return $sheet->{active};
    } # label

# my @row = $sheet->cellrow (1);
sub cellrow {
    my ($sheet, $row) = @_;
    defined $row && $row > 0 && $row <= $sheet->{maxrow} or return;
    my $s = $sheet->{cell};
    map { $s->[$_][$row] } 1..$sheet->{maxcol};
    } # cellrow

# my @row = $sheet->row (1);
sub row {
    my ($sheet, $row) = @_;
    defined $row && $row > 0 && $row <= $sheet->{maxrow} or return;
    map { $sheet->{$sheet->cr2cell ($_, $row)} } 1..$sheet->{maxcol};
    } # row

# my @col = $sheet->cellcolumn (1);
sub cellcolumn {
    my ($sheet, $col) = @_;
    defined $col && $col > 0 && $col <= $sheet->{maxcol} or return;
    my $s = $sheet->{cell};
    map { $s->[$col][$_] } 1..$sheet->{maxrow};
    } # cellcolumn

# my @col = $sheet->column (1);
sub column {
    my ($sheet, $col) = @_;
    defined $col && $col > 0 && $col <= $sheet->{maxcol} or return;
    map { $sheet->{$sheet->cr2cell ($col, $_)} } 1..$sheet->{maxrow};
    } # column

# Convert {cell}'s [column][row] to a [row][column] list
# my @rows = $sheet->rows ();
sub rows {
    my $sheet = shift;
    my $s = $sheet->{cell};

    map {
	my $r = $_;
	[ map { $s->[$_][$r] } 1..$sheet->{maxcol} ];
	} 1..$sheet->{maxrow};
    } # rows

1;

__END__
=head1 DESCRIPTION

Spreadsheet::Read tries to transparently read *any* spreadsheet and
return its content in a universal manner independent of the parsing
module that does the actual spreadsheet scanning.

For OpenOffice and/or LibreOffice this module uses
L<Spreadsheet::ReadSXC|https://metacpan.org/release/Spreadsheet-ReadSXC>

For Microsoft Excel this module uses
L<Spreadsheet::ParseExcel|https://metacpan.org/release/Spreadsheet-ParseExcel>,
L<Spreadsheet::ParseXLSX|https://metacpan.org/release/Spreadsheet-ParseXLSX>, or
L<Spreadsheet::XLSX|https://metacpan.org/release/Spreadsheet-XLSX> (stronly
discouraged).

For CSV this module uses L<Text::CSV_XS|https://metacpan.org/release/Text-CSV_XS>
or L<Text::CSV_PP|https://metacpan.org/release/Text-CSV_PP>.

For SquirrelCalc there is a very simplistic built-in parser

=head2 Data structure

The data is returned as an array reference:

  $book = [
      # Entry 0 is the overall control hash
      { sheets  => 2,
        sheet   => {
          "Sheet 1"  => 1,
          "Sheet 2"  => 2,
          },
        parsers => [ {
            type    => "xls",
            parser  => "Spreadsheet::ParseExcel",
            version => 0.59,
            }],
        error   => undef,
        },
      # Entry 1 is the first sheet
      { parser  => 0,
        label   => "Sheet 1",
        maxrow  => 2,
        maxcol  => 4,
        cell    => [ undef,
          [ undef, 1 ],
          [ undef, undef, undef, undef, undef, "Nugget" ],
          ],
        attr    => [],
        merged  => [],
        active  => 1,
        A1      => 1,
        B5      => "Nugget",
        },
      # Entry 2 is the second sheet
      { parser  => 0,
        label   => "Sheet 2",
        :
        :

To keep as close contact to spreadsheet users, row and column 1 have
index 1 too in the C<cell> element of the sheet hash, so cell "A1" is
the same as C<cell> [1, 1] (column first). To switch between the two,
there are helper functions available: C<cell2cr ()>, C<cr2cell ()>,
and C<col2label ()>.

The C<cell> hash entry contains unformatted data, while the hash entries
with the traditional labels contain the formatted values (if applicable).

The control hash (the first entry in the returned array ref), contains
some spreadsheet meta-data. The entry C<sheet> is there to be able to find
the sheets when accessing them by name:

  my %sheet2 = %{$book->[$book->[0]{sheet}{"Sheet 2"}]};

=head2 Functions and methods

=head3 new

 my $book = Spreadsheet::Read->new (...);

All options accepted by ReadData are accepted by new.

=head3 ReadData

 my $book = ReadData ($source [, option => value [, ... ]]);

 my $book = ReadData ("file.csv", sep => ',', quote => '"');

 my $book = ReadData ("file.xls", dtfmt => "yyyy-mm-dd");

 my $book = ReadData ("file.ods");

 my $book = ReadData ("file.sxc");

 my $book = ReadData ("content.xml");

 my $book = ReadData ($content);

 my $book = ReadData ($content,  parser => "xlsx");

 my $book = ReadData ($fh,       parser => "xlsx");

 my $book = ReadData (\$content, parser => "xlsx");

Tries to convert the given file, string, or stream to the data structure
described above.

Processing Excel data from a stream or content is supported through a
L<File::Temp|https://metacpan.org/release/File-Temp> temporary file or
L<IO::Scalar|https://metacpan.org/release/IO-Scalar> when available.

L<Spreadsheet::ReadSXC|https://metacpan.org/release/Spreadsheet-ReadSXC>
does preserve sheet order as of version 0.20.

Choosing between C<$content> and C<\\$content> (with or without passing
the desired C<parser> option) may be depending on trial and terror.
C<ReadData> does try to determine parser type on content if needed, but
not all combinations are checked, and not all signatures are builtin.

Currently supported options are:

=over 2

=item parser
X<parser>

Force the data to be parsed by a specific format. Possible values are
C<csv>, C<prl> (or C<perl>), C<sc> (or C<squirelcalc>), C<sxc> (or C<oo>,
C<ods>, C<openoffice>, C<libreoffice>) C<xls> (or C<excel>), and C<xlsx>
(or C<excel2007>).

When parsing streams, instead of files, it is highly recommended to pass
this option.

Spreadsheet::Read supports several underlying parsers per spreadsheet
type. It will try those from most favored to least favored. When you
have a good reason to prefer a different parser, you can set that in
environment variables. The other options then will not be tested for:

 env SPREADSHEET_READ_CSV=Text::CSV_PP ...

=item cells
X<cells>

Control the generation of named cells ("C<A1>" etc). Default is true.

=item rc

Control the generation of the {cell}[c][r] entries. Default is true.

=item attr

Control the generation of the {attr}[c][r] entries. Default is false.
See L</Cell Attributes> below.

=item clip

If set, L<C<ReadData>|/ReadData> will remove all trailing rows and columns
per sheet that have no data, where no data means only undefined or empty
cells (after optional stripping). If a sheet has no data at all, the sheet
will be skipped entirely when this attribute is true.

=item strip

If set, L<C<ReadData>|/ReadData> will remove trailing- and/or
leading-whitespace from every field.

  strip  leading  strailing
  -----  -------  ---------
    0      n/a      n/a
    1     strip     n/a
    2      n/a     strip
    3     strip    strip

=item pivot

Swap all rows and columns.

When a sheet contains data like

  A1  B1  C1      E1
  A2      C2  D2
  A3  B3  C3  D3  E3

using C<pivot> will return the sheet data as

  A1  A2  A3
  B1      B3
  C1  C2  C3
      D2  D3
  E1      E3

=item sep

Set separator for CSV. Default is comma C<,>.

=item quote

Set quote character for CSV. Default is C<">.

=item dtfmt

Set the format for MS-Excel date fields that are set to use the default
date format. The default format in Excel is "C<m-d-yy>", which is both
not year 2000 safe, nor very useful. The default is now "C<yyyy-mm-dd>",
which is more ISO-like.

Note that date formatting in MS-Excel is not reliable at all, as it will
store/replace/change the date field separator in already stored formats
if you change your locale settings. So the above mentioned default can
be either "C<m-d-yy>" OR "C<m/d/yy>" depending on what that specific
character happened to be at the time the user saved the file.

=item debug

Enable some diagnostic messages to STDERR.

The value determines how much diagnostics are dumped (using
L<Data::Peek|https://metacpan.org/release/Data-Peek>).  A value of C<9>
and higher will dump the entire structure from the back-end parser.

=item passwd

Use this password to decrypt password protected spreadsheet.

Currently only supports Excel.

=back

All other attributes/options will be passed to the underlying parser if
that parser supports attributes.

=head3 col2label

 my $col_id = col2label (col);

 my $col_id = $book->col2label (col);  # OO

C<col2label ()> converts a C<(column)> (1 based) to the letters used in the
traditional cell notation:

  my $id = col2label ( 4); # $id now "D"
  my $id = col2label (28); # $id now "AB"

=head3 cr2cell

 my $cell = cr2cell (col, row);

 my $cell = $book->cr2cell (col, row);  # OO

C<cr2cell ()> converts a C<(column, row)> pair (1 based) to the
traditional cell notation:

  my $cell = cr2cell ( 4, 14); # $cell now "D14"
  my $cell = cr2cell (28,  4); # $cell now "AB4"

=head3 cell2cr

 my ($col, $row) = cell2cr ($cell);

 my ($col, $row) = $book->cell2cr ($cell);  # OO

C<cell2cr ()> converts traditional cell notation to a C<(column, row)>
pair (1 based):

  my ($col, $row) = cell2cr ("D14"); # returns ( 4, 14)
  my ($col, $row) = cell2cr ("AB4"); # returns (28,  4)

=head3 row

 my @row = row ($sheet, $row)

 my @row = Spreadsheet::Read::row ($book->[1], 3);

 my @row = $book->row ($sheet, $row); # OO

Get full row of formatted values (like C<< $sheet->{A3} .. $sheet->{G3} >>)

Note that the indexes in the returned list are 0-based.

C<row ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=head3 cellrow

 my @row = cellrow ($sheet, $row);

 my @row = Spreadsheet::Read::cellrow ($book->[1], 3);

 my @row = $book->cellrow ($sheet, $row); # OO

Get full row of unformatted values (like C<< $sheet->{cell}[1][3] .. $sheet->{cell}[7][3] >>)

Note that the indexes in the returned list are 0-based.

C<cellrow ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified or as method call.

=head3 rows

 my @rows = rows ($sheet);

 my @rows = Spreadsheet::Read::rows ($book->[1]);

 my @rows = $book->rows (1); # OO

Convert C<{cell}>'s C<[column][row]> to a C<[row][column]> list.

Note that the indexes in the returned list are 0-based, where the
index in the C<{cell}> entry is 1-based.

C<rows ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=head3 parses

 parses ($format);

 Spreadsheet::Read::parses ("CSV");

 $book->parses ("CSV"); # OO

C<parses ()> returns Spreadsheet::Read's capability to parse the
required format. L<C<ReadData>|/ReadData> will pick its preferred parser
for that format unless overruled. See L<C<parser>|/parser>.

C<parses ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=head3 Version

 my $v = Version ()

 my $v = Spreadsheet::Read::Version ()

 my $v = Spreadsheet::Read->VERSION;

 my $v = $book->Version (); # OO

Returns the current version of Spreadsheet::Read.

C<Version ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

This function returns exactly the same as C<< Spreadsheet::Read->VERSION >>
returns and is only kept for backward compatibility reasons.

=head3 sheets

 my $sheets = $book->sheets; # OO
 my @sheets = $book->sheets; # OO

In scalar context return the number of sheets in the book.
In list context return the labels of the sheets in the book.

=head3 sheet

 my $sheet = $book->sheet (1);     # OO
 my $sheet = $book->sheet ("Foo"); # OO

Return the numbered or named sheet out of the book. Will return C<undef> if
there is no match. Will not work for sheets I<named> with a number between 1
and the number of sheets in the book.

With named sheets will first try to use the list of sheet-labels as stored in
the control structure. If no match is found, it will scan the actual labels
of the sheets. In that case, it will return the first matching sheet.

If defined, the returned sheet will be of class C<Spreadsheet::Read::Sheet>.

=head3 add

 my $book = ReadData ("file.csv");
 Spreadsheet::Read::add ($book, "file.xlsx");

 my $book = Spreadsheet::Read->new ("file.csv");
 $book->add ("file.xlsx"); # OO

=head2 Methods on sheets

=head3 maxcol

 my $col = $sheet->maxcol;

Return the index of the last in-use column in the sheet. This index is 1-based.

=head3 maxrow

 my $row = $sheet->maxrow;

Return the index of the last in-use row in the sheet. This index is 1-based.

=head3 cell

 my $cell = $sheet->cell ("A3");
 my $cell = $sheet->cell (1, 3);

Return the value for a cell. Using tags will return the formatted value,
using column and row will return unformatted value.

=head3 attr

 my $cell = $sheet->attr ("A3");
 my $cell = $sheet->attr (1, 3);

Return the attributes of a cell. Only valid if attributes are enabled through
option C<attr>.

=head3 col2label

 my $col_id = $sheet->col2label (col);

C<col2label ()> converts a C<(column)> (1 based) to the letters used in the
traditional cell notation:

  my $id = $sheet->col2label ( 4); # $id now "D"
  my $id = $sheet->col2label (28); # $id now "AB"

=head3 cr2cell

 my $cell = $sheet->cr2cell (col, row);

C<cr2cell ()> converts a C<(column, row)> pair (1 based) to the
traditional cell notation:

  my $cell = $sheet->cr2cell ( 4, 14); # $cell now "D14"
  my $cell = $sheet->cr2cell (28,  4); # $cell now "AB4"

=head3 cell2cr

 my ($col, $row) = $sheet->cell2cr ($cell);

C<cell2cr ()> converts traditional cell notation to a C<(column, row)>
pair (1 based):

  my ($col, $row) = $sheet->cell2cr ("D14"); # returns ( 4, 14)
  my ($col, $row) = $sheet->cell2cr ("AB4"); # returns (28,  4)

=head3 col

 my @col = $sheet->column ($col);

Get full column of formatted values (like C<< $sheet->{C1} .. $sheet->{C9} >>)

Note that the indexes in the returned list are 0-based.

=head3 cellcolumn

 my @col = $sheet->cellcolumn ($col);

Get full column of unformatted values (like C<< $sheet->{cell}[3][1] .. $sheet->{cell}[3][9] >>)

Note that the indexes in the returned list are 0-based.

=head3 row

 my @row = $sheet->row ($row);

Get full row of formatted values (like C<< $sheet->{A3} .. $sheet->{G3} >>)

Note that the indexes in the returned list are 0-based.

=head3 cellrow

 my @row = $sheet->cellrow ($row);

Get full row of unformatted values (like C<< $sheet->{cell}[1][3] .. $sheet->{cell}[7][3] >>)

Note that the indexes in the returned list are 0-based.

=head3 rows

 my @rows = $sheet->rows ();

Convert C<{cell}>'s C<[column][row]> to a C<[row][column]> list.

Note that the indexes in the returned list are 0-based, where the
index in the C<{cell}> entry is 1-based.

=head3 label

 my $label = $sheet->label;
 $sheet->label ("New sheet label");

Set a new label to a sheet. Note that the index in the control structure will
I<NOT> be updated.

=head3 active

 my $sheet_is_active = $sheet->active;

Returns 1 if the selected sheet is active, otherwise returns 0.

Currently only works on XLS (as of Spreadsheed::ParseExcel-0.61).
CSV is always active.

=head2 Using CSV

In case of CSV parsing, L<C<ReadData>|/ReadData> will use the first line of
the file to auto-detect the separation character if the first argument is a
file and both C<sep> and C<quote> are not passed as attributes.
L<Text::CSV_XS|https://metacpan.org/release/Text-CSV_XS> (or
L<Text::CSV_PP|https://metacpan.org/release/Text-CSV_PP>) is able to
automatically detect and use C<\r> line endings.

CSV can parse streams too, but be sure to pass C<sep> and/or C<quote> if
these do not match the default C<,> and C<">.

When an error is found in the CSV, it is automatically reported (to STDERR).
The structure will store the error in C<< $ss->[0]{error} >> as anonymous
list returned by
L<C<< $csv->error_diag >>|https://metacpan.org/pod/Text::CSV_XS#error_diag>.
See L<Text::CSV_XS|https://metacpan.org/pod/Text-CSV_XS> for documentation.

 my $ss = ReadData ("bad.csv");
 $ss->[0]{error} and say $ss->[0]{error}[1];

As CSV has no sheet labels, the default label for a CSV sheet is its filename.
For CSV, this can be overruled using the I<label> attribute:

 my $ss = Spreadsheet::Read->new ("/some/place/test.csv", label => "Test");

=head2 Cell Attributes
X<attr>

If the constructor was called with C<attr> having a true value,

 my $book = ReadData ("book.xls", attr => 1);
 my $book = Spreadsheet::Read->new ("book.xlsx", attr => 1);

effort is made to analyze and store field attributes like this:

    { label  => "Sheet 1",
      maxrow => 5,
      maxcol => 2,
      cell   => [ undef,
	[ undef, 1 ],
	[ undef, undef, undef, undef, undef, "Nugget" ],
	],
      attr   => [ undef,
	[ undef, {
	  type    => "numeric",
	  fgcolor => "#ff0000",
	  bgcolor => undef,
	  font    => "Arial",
	  size    => undef,
	  format  => "## ##0.00",
	  halign  => "right",
	  valign  => "top",
	  uline   => 0,
	  bold    => 0,
	  italic  => 0,
	  wrap    => 0,
	  merged  => 0,
	  hidden  => 0,
	  locked  => 0,
	  enc     => "utf-8",
	  }, ]
	[ undef, undef, undef, undef, undef, {
	  type    => "text",
	  fgcolor => "#e2e2e2",
	  bgcolor => undef,
	  font    => "Letter Gothic",
	  size    => 15,
	  format  => undef,
	  halign  => "left",
	  valign  => "top",
	  uline   => 0,
	  bold    => 0,
	  italic  => 0,
	  wrap    => 0,
	  merged  => 0,
	  hidden  => 0,
	  locked  => 0,
	  enc     => "iso8859-1",
	  }, ]
      merged => [],
      A1     => 1,
      B5     => "Nugget",
      },

The entries C<maxrow> and C<maxcol> are 1-based.

This has now been partially implemented, mainly for Excel, as the other
parsers do not (yet) support all of that. YMMV.

If a cell itself is not hidden, but the parser holds the information that
either the row or the column (or both) the field is in is hidden, the flag
is inherited into the cell attributes.

You can get the attributes of a cell (as a hash-ref) like this:

 my $attr = $book[1]{attr}[1][3];          # Direct structure
 my $attr = $book->sheet (1)->attr (1, 3); # Same using OO
 my $attr = $book->sheet (1)->attr ("A3"); # Same using OO

To get to the C<font> attribute, use any of these:

 my $font = $book[1]{attr}[1][3]{font};
 my $font = $book->sheet (1)->attr (1, 3)->{font};
 my $font = $book->sheet (1)->attr ("A3")->font;

=head3 Merged cells
X<merged>

Note that only
L<Spreadsheet::ReadSXC|https://metacpan.org/release/Spreadsheet-ReadSXC>
documents the use of merged cells, and not in a way useful for the spreadsheet
consumer.

CSV does not support merged cells (though future implementations of CSV
for the web might).

The documentation of merged areas in
L<Spreadsheet::ParseExcel|https://metacpan.org/release/Spreadsheet-ParseExcel> and
L<Spreadsheet::ParseXLSX|https://metacpan.org/release/Spreadsheet-ParseXLSX> can
be found in
L<Spreadsheet::ParseExcel::Worksheet|https://metacpan.org/release/Spreadsheet-ParseExcel-Worksheet>
and L<Spreadsheet::ParseExcel::Cell|https://metacpan.org/release/Spreadsheet-ParseExcel-Cell>.

None of basic L<Spreadsheet::XLSX|https://metacpan.org/release/Spreadsheet-XLSX>,
L<Spreadsheet::ParseExcel|https://metacpan.org/release/Spreadsheet-ParseExcel>, and
L<Spreadsheet::ParseXLSX|https://metacpan.org/release/Spreadsheet-ParseXLSX> manual
pages mention merged cells at all.

This module just tries to return the information in a generic way.

Given this spreadsheet as an example

 merged.xlsx:

     A     B     C
  +-----+-----------+
 1|     | foo       |
  +-----+           +
 2| bar |           |
  |     +-----+-----+
 3|     | urg | orc |
  +-----+-----+-----+

the information extracted from that undocumented information is
returned in the C<merged> entry of the sheet's hash as a list of
top-left, bottom-right coordinate pars (col, row, col, row). For
given example, that would be:

 $ss->{merged} = [
    [ 1, 2, 1, 3 ], # A2-A3
    [ 2, 1, 3, 2 ], # B1-C2
    ];

When the attributes are also enabled, there is some merge information
copied directly from the cell information, but again, that stems from
code analysis and not from documentation:

 my $ss = ReadData ("merged.xlsx", attr => 1)->[1];
 foreach my $row (1 .. $ss->{maxrow}) {
     foreach my $col (1 .. $ss->{maxcol}) {
         my $cell = cr2cell ($col, $row);
         printf "%s %-3s %d  ", $cell, $ss->{$cell},
             $ss->{attr}[$col][$row]{merged};
         }
     print "\n";
     }

 A1     0  B1 foo 1  C1     1
 A2 bar 1  B2     1  C2     1
 A3     1  B3 urg 0  C3 orc 0

In this example, there is no way to see if C<B2> is merged to C<A2> or
to C<B1> without analyzing all surrounding cells. This could as well
mean C<A2:A3>, C<B1:C1>, C<B2:C2>, as C<A2:A3>, C<B1:B2>, C<C1:C2>, as
C<A2:A3>, C<B1:C2>.
Use the L<C<merged>|/merged> entry described above to find out what
fields are merged to what other fields.

=head1 TOOLS

This modules comes with a few tools that perform tasks from the FAQ, like
"How do I select only column D through F from sheet 2 into a CSV file?"

If the module was installed without the tools, you can find them here:
  https://github.com/Tux/Spreadsheet-Read/tree/master/examples

=head2 C<xlscat>

Show (parts of) a spreadsheet in plain text, CSV, or HTML

 usage: xlscat   [-s <sep>] [-L] [-n] [-A] [-u] [Selection] file.xls
                 [-c | -m]                 [-u] [Selection] file.xls
                  -i                            [-S sheets] file.xls
    Generic options:
       -v[#]       Set verbose level (xlscat/xlsgrep)
       -d[#]       Set debug   level (Spreadsheet::Read)
       -u          Use unformatted values
       --noclip    Do not strip empty sheets and
                   trailing empty rows and columns
       -e <enc>    Set encoding for input and output
       -b <enc>    Set encoding for input
       -a <enc>    Set encoding for output
    Input CSV:
       --in-sep=c  Set input sep_char for CSV
    Input XLS:
       --dtfmt=fmt Specify the default date format to replace 'm-d-yy'
                   the default replacement is 'yyyy-mm-dd'
    Output Text (default):
       -s <sep>    Use separator <sep>. Default '|', \n allowed
       -L          Line up the columns
       -n [skip]   Number lines (prefix with column number)
                   optionally skip <skip> (header) lines
       -A          Show field attributes in ANSI escapes
       -h[#]       Show # header lines
    Output Index only:
       -i          Show sheet names and size only
    Output CSV:
       -c          Output CSV, separator = ','
       -m          Output CSV, separator = ';'
    Output HTML:
       -H          Output HTML
    Selection:
       -S <sheets> Only print sheets <sheets>. 'all' is a valid set
                   Default only prints the first sheet
       -R <rows>   Only print rows    <rows>. Default is 'all'
       -C <cols>   Only print columns <cols>. Default is 'all'
       -F <flds>   Only fields <flds> e.g. -FA3,B16
    Ordering (column numbers in result set *after* selection):
       --sort=spec Sort output (e.g. --sort=3,2r,5n,1rn+2)
                   +#   - first # lines do not sort (header)
                   #    - order on column # lexical ascending
                   #n   - order on column # numeric ascending
                   #r   - order on column # lexical descending
                   #rn  - order on column # numeric descending

=head2 C<xlsgrep>

Show (parts of) a spreadsheet that match a pattern in plain text, CSV, or HTML

 usage: xlsgrep  [-s <sep>] [-L] [-n] [-A] [-u] [Selection] pattern file.xls
                 [-c | -m]                 [-u] [Selection] pattern file.xls
                  -i                            [-S sheets] pattern file.xls
    Generic options:
       -v[#]       Set verbose level (xlscat/xlsgrep)
       -d[#]       Set debug   level (Spreadsheet::Read)
       -u          Use unformatted values
       --noclip    Do not strip empty sheets and
                   trailing empty rows and columns
       -e <enc>    Set encoding for input and output
       -b <enc>    Set encoding for input
       -a <enc>    Set encoding for output
    Input CSV:
       --in-sep=c  Set input sep_char for CSV
    Input XLS:
       --dtfmt=fmt Specify the default date format to replace 'm-d-yy'
                   the default replacement is 'yyyy-mm-dd'
    Output Text (default):
       -s <sep>    Use separator <sep>. Default '|', \n allowed
       -L          Line up the columns
       -n [skip]   Number lines (prefix with column number)
                   optionally skip <skip> (header) lines
       -A          Show field attributes in ANSI escapes
       -h[#]       Show # header lines
    Grep options:
       -i          Ignore case
       -w          Match whole words only
    Output CSV:
       -c          Output CSV, separator = ','
       -m          Output CSV, separator = ';'
    Output HTML:
       -H          Output HTML
    Selection:
       -S <sheets> Only print sheets <sheets>. 'all' is a valid set
                   Default only prints the first sheet
       -R <rows>   Only print rows    <rows>. Default is 'all'
       -C <cols>   Only print columns <cols>. Default is 'all'
       -F <flds>   Only fields <flds> e.g. -FA3,B16
    Ordering (column numbers in result set *after* selection):
       --sort=spec Sort output (e.g. --sort=3,2r,5n,1rn+2)
                   +#   - first # lines do not sort (header)
                   #    - order on column # lexical ascending
                   #n   - order on column # numeric ascending
                   #r   - order on column # lexical descending
                   #rn  - order on column # numeric descending

=head2 C<xls2csv>

Convert a spreadsheet to CSV. This is just a small wrapper over C<xlscat>.

 usage: xls2csv [ -o file.csv ] file.xls

=head2 C<ss2tk>

Show a spreadsheet in a perl/Tk spreadsheet widget

 usage: ss2tk [-w <width>] [X11 options] file.xls [<pattern>]
        -w <width> use <width> as default column width (4)

=head2 C<ssdiff>

Show the differences between two spreadsheets.

 usage: examples/ssdiff [--verbose[=1]] file.xls file.xlsx

=head1 TODO

=over 4

=item Options

=over 2

=item Module Options

New Spreadsheet::Read options are bound to happen. I'm thinking of an
option that disables the reading of the data entirely to speed up an
index request (how many sheets/fields/columns). See C<xlscat -i>.

=item Parser options

Try to transparently support as many options as the encapsulated modules
support regarding (un)formatted values, (date) formats, hidden columns
rows or fields etc. These could be implemented like C<attr> above but
names C<meta>, or just be new values in the C<attr> hashes.

=back

=item Other parsers

Add support for new(er) parsers for already supported formats, like

=over 2

=item Data::XLSX::Parser

Data::XLSX::Parser provides faster way to parse Microsoft Excel's .xlsx
files. The implementation of this module is highly inspired from Python's
FastXLSX library.

This is SAX based parser, so you can parse very large XLSX file with
lower memory usage.

=back

=item Other spreadsheet formats

I consider adding any spreadsheet interface that offers a usable API.

Under investigation:

=over 2

=item Gnumeric (.gnumeric)

I have seen no existing CPAN module yet.

It is gzip'ed XML

=item Kspread (.ksp)

Now knows as Calligra Sheets.

I have seen no existing CPAN module yet.

It is XML in ZIP

=back

=item Alternative parsers for existing formats

As long as the alternative has a good reason for its existence, and the
API of that parser reasonable fits in my approach, I will consider to
implement the glue layer, or apply patches to do so as long as these
match what F<CONTRIBUTING.md> describes.

=back

=head1 SEE ALSO

=over 2

=item Text::CSV_XS, Text::CSV_PP

See L<Text::CSV_XS|https://metacpan.org/release/Text-CSV_XS> ,
L<Text::CSV_PP|https://metacpan.org/release/Text-CSV_PP> , and
L<Text::CSV|https://metacpan.org/release/Text-CSV> documentation.

L<Text::CSV|https://metacpan.org/release/Text-CSV> is a wrapper over Text::CSV_XS (the fast XS version) and/or
L<Text::CSV_PP|https://metacpan.org/release/Text-CSV_PP> (the pure perl version).

=item Spreadsheet::ParseExcel

L<Spreadsheet::ParseExcel|https://metacpan.org/release/Spreadsheet-ParseExcel> is
the best parser for old-style Microsoft Excel (.xls) files.

=item Spreadsheet::ParseXLSX

L<Spreadsheet::ParseXLSX|https://metacpan.org/release/Spreadsheet-ParseXLSX> is
like L<Spreadsheet::ParseExcel|https://metacpan.org/release/Spreadsheet-ParseExcel>,
but for new Microsoft Excel 2007+ files (.xlsx). They have the same API.

This module uses L<XML::Twig|https://metacpan.org/release/XML-Twig> to parse the
internal XML.

=item Spreadsheet::XLSX

See L<Spreadsheet::XLSX|https://metacpan.org/release/Spreadsheet-XLSX>
documentation.

This module is dead and deprecated. It is B<buggy and unmaintained>.  I<Please>
use L<Spreadsheet::ParseXLSX|https://metacpan.org/release/Spreadsheet-ParseXLSX>
instead.

=item Spreadsheet::ReadSXC

L<Spreadsheet::ReadSXC|https://metacpan.org/release/Spreadsheet-ReadSXC> is a
parser for OpenOffice/LibreOffice (.sxc and .ods) spreadsheet files.

=item Spreadsheet::BasicRead

See L<Spreadsheet::BasicRead|https://metacpan.org/release/Spreadsheet-BasicRead>
for xlscat-like functionality (Excel only)

=item Spreadsheet::ConvertAA

See L<Spreadsheet::ConvertAA|https://metacpan.org/release/Spreadsheet-ConvertAA>
for an alternative set of L</cell2cr>/L</cr2cell> pair.

=item Spreadsheet::Perl

L<Spreadsheet::Perl|https://metacpan.org/release/Spreadsheet-Perl> offers a Pure
Perl implementation of a spreadsheet engine.  Users that want this format to be
supported in Spreadsheet::Read are hereby motivated to offer patches. It is
not high on my TODO-list.

=item Spreadsheet::CSV

L<Spreadsheet::CSV|https://metacpan.org/release/Spreadsheet-CSV> offers the
interesting approach of seeing all supported spreadsheet formats as if it were
CSV, mimicking the L<Text::CSV_XS|https://metacpan.org/release/Text-CSV_XS>
interface.

=item xls2csv

L<xls2csv|https://metacpan.org/release/xls2csv> offers an alternative for my
C<xlscat -c>, in the xls2csv tool, but this tool focuses on character encoding
transparency, and requires some other modules.

=back

=head1 AUTHOR

H.Merijn Brand, <h.m.brand@xs4all.nl>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2005-2018 H.Merijn Brand

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
