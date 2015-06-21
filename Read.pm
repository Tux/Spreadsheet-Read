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

 my $sheet = $book->[1];             # first datasheet
 my $cell  = $book->[1]{A3};         # content of field A3 of sheet 1
 my $cell  = $book->[1]{cell}[1][3]; # same, unformatted

=cut

use strict;
use warnings;

our $VERSION = "0.62";
sub  Version { $VERSION }

use Carp;
use Exporter;
our @ISA       = qw( Exporter );
our @EXPORT    = qw( ReadData cell2cr cr2cell );
our @EXPORT_OK = qw( parses rows cellrow row );

use File::Temp   qw( );
use Data::Dumper;

my @parsers = (
    [ csv  => "Text::CSV_XS",				"0.71"	],
    [ csv  => "Text::CSV_PP",				"1.17"	],
    [ csv  => "Text::CSV",				"1.17"	],
    [ ods  => "Spreadsheet::ReadSXC",			"0.20"	],
    [ sxc  => "Spreadsheet::ReadSXC",			"0.20"	],
    [ xls  => "Spreadsheet::ParseExcel",		"0.34"	],
    [ xlsx => "Spreadsheet::ParseXLSX",			"0.13"	],
    [ xlsx => "Spreadsheet::XLSX",			"0.13"	],
    [ xlsx => "Spreadsheet::XLSX::Reader::LibXML",	"v0.36"	],
    [ prl  => "Spreadsheet::Perl",			""	],

    # Helper modules
    [ ios  => "IO::Scalar",				""	],
    [ dmp  => "Data::Peek",				""	],
    );
my %can = map {
    my $preset = $ENV{"SPREADSHEET_READ_\U$_->[0]"};
    $preset and eval "require $preset";
    $_->[0] => $preset;
    } @parsers;
for (@parsers) {
    my ($flag, $mod, $vsn) = @$_;
    $can{$flag} and next;
    eval "require $mod; \$vsn and ${mod}->VERSION (\$vsn); \$can{\$flag} = '$mod'";
    #$@ && $@ !~ m{Can't locate} and warn "$mod: $@\n";
    }
$can{sc} = __PACKAGE__;	# SquirelCalc is built-in

my $debug = 0;
my %def_opts = (
    rc      => 1,
    cells   => 1,
    attr    => 0,
    clip    => undef, # $opt{cells};
    strip   => 0,
    dtfmt   => "yyyy-mm-dd", # Format 14
    debug   => 0,
    parser  => undef,
    sep     => undef,
    quote   => undef,
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
    );

# Helper functions

sub _dump
{
    my ($label, $ref) = @_;
    if ($can{dmp}) {
	print STDERR Data::Peek::DDumper ({ $label => $ref });
	}
    else {
	print STDERR Data::Dumper->Dump ([$ref], [$label]);
	}
    } # _dump

sub _parser
{
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

# Spreadsheet::Read::parses ("csv") or die "Cannot parse CSV"
sub parses
{
    my $type = _parser (shift)	or  return 0;
    return $can{$type};
    } # parses

# cr2cell (4, 18) => "D18"
# No prototype to allow 'cr2cell (@rowcol)'
sub cr2cell
{
    my ($c, $r) = @_;
    defined $c && defined $r && $c > 0 && $r > 0 or return "";
    my $cell = "";
    while ($c) {
	use integer;

	substr $cell, 0, 0, chr (--$c % 26 + ord "A");
	$c /= 26;
	}
    "$cell$r";
    } # cr2cell

# cell2cr ("D18") => (4, 18)
sub cell2cr
{
    my ($cc, $r) = (uc ($_[0]||"") =~ m/^([A-Z]+)([0-9]+)$/) or return (0, 0);
    my $c = 0;
    while ($cc =~ s/^([A-Z])//) {
	$c = 26 * $c + 1 + ord ($1) - ord ("A");
	}
    ($c, $r);
    } # cell2cr

# my @row = cellrow ($book->[1], 1);
sub cellrow
{
    my $sheet = shift or return;
    ref     $sheet eq "HASH" && exists  $sheet->{cell}   or return;
    exists  $sheet->{maxcol} && exists  $sheet->{maxrow} or return;
    my $row   = shift or return;
    $row > 0 && $row <= $sheet->{maxrow} or return;
    my $s = $sheet->{cell};
    map { $s->[$_][$row] } 1..$sheet->{maxcol};
    } # cellrow

# my @row = row ($book->[1], 1);
sub row
{
    my $sheet = shift or return;
    ref     $sheet eq "HASH" && exists  $sheet->{cell}   or return;
    exists  $sheet->{maxcol} && exists  $sheet->{maxrow} or return;
    my $row   = shift or return;
    $row > 0 && $row <= $sheet->{maxrow} or return;
    map { $sheet->{cr2cell ($_, $row)} } 1..$sheet->{maxcol};
    } # row

# Convert {cell}'s [column][row] to a [row][column] list
# my @rows = rows ($book->[1]);
sub rows
{
    my $sheet = shift or return;
    ref    $sheet eq "HASH" && exists $sheet->{cell}   or return;
    exists $sheet->{maxcol} && exists $sheet->{maxrow} or return;
    my $s = $sheet->{cell};

    map {
	my $r = $_;
	[ map { $s->[$_][$r] } 1..$sheet->{maxcol} ];
	} 1..$sheet->{maxrow};
    } # rows

# If option "clip" is set, remove the trailing rows and
# columns in each sheet that contain no visible data
sub _clipsheets
{
    my ($opt, $ref) = @_;

    if (my $s = $opt->{strip} and $ref->[0]{sheets}) {
	foreach my $sheet (1 .. $ref->[0]{sheets}) {
	    my $ss = $ref->[$sheet];
	    $ss->{maxrow} && $ss->{maxcol} or next;
	    foreach my $row (1 .. $ss->{maxrow}) {
		foreach my $col (1 .. $ss->{maxcol}) {
		    for (($opt->{rc}    ? $ss->{cell}   [$col][$row]  : ()),
		         ($opt->{cells} ? $ss->{cr2cell ($col, $row)} : ())) {
			defined or next;
		        $s & 2 and s/\s+$//;
			$s & 1 and s/^\s+//;
			}
		    }
		}
	    }
	}

    $opt->{clip} or return $ref;

    foreach my $sheet (1 .. $ref->[0]{sheets}) {
	my $ss = $ref->[$sheet];

	# Remove trailing empty columns
	while ($ss->{maxcol} and not (
		grep { defined && m/\S/ } @{$ss->{cell}[$ss->{maxcol}]})
		) {
	    (my $col = cr2cell ($ss->{maxcol}, 1)) =~ s/1$//;
	    my $recol = qr{^$col(?=[0-9]+)$};
	    delete $ss->{$_} for grep m/$recol/, keys %{$ss};
	    $ss->{maxcol}--;
	    }
	$ss->{maxcol} or $ss->{maxrow} = 0;

	# Remove trailing empty rows
	while ($ss->{maxrow} and not (
		grep { defined && m/\S/ }
		map  { $ss->{cell}[$_][$ss->{maxrow}] }
		1 .. $ss->{maxcol}
		)) {
	    my $rerow = qr{^[A-Z]+$ss->{maxrow}$};
	    delete $ss->{$_} for grep m/$rerow/, keys %{$ss};
	    $ss->{maxrow}--;
	    }
	$ss->{maxrow} or $ss->{maxcol} = 0;
	}
    $ref;
    } # _clipsheets

# Convert a single color (index) to a color
sub _xls_color
{
    my $clr = shift;
    defined $clr		or  return undef;
    $clr eq "#000000"		and return undef;
    $clr =~ m/^#[0-9a-fA-F]+$/	and return lc $clr;
    $clr == 0 || $clr == 32767	and return undef; # Default fg color
    return "#" . lc Spreadsheet::ParseExcel->ColorIdxToRGB ($clr);
    } # _xls_color

# Convert a fill [ $pattern, $front_color, $back_color ] to a single background
sub _xls_fill
{
    my ($p, $fg, $bg) = @_;
    defined $p			or  return undef;
    $p == 32767			and return undef; # Default fg color
    $p == 0 && !defined $bg	and return undef; # No fill bg color
    $p == 1			and return _xls_color ($fg);
    $bg < 8 || $bg > 63		and return undef; # see Workbook.pm#106
    return _xls_color ($bg);
    } # _xls_fill

sub ReadData
{
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

    my $io_ref = ref ($txt) =~ m/GLOB|IO/ ? $txt : undef;
    my $io_fil = $io_ref ? 0 : do { no warnings "newline"; -f $txt ? 1 : 0 };
    my $io_txt = $io_ref || $io_fil ? 0 : 1;

    $io_fil && ! -s $txt  and return;
    $io_ref && eof ($txt) and return;

    if ($opt{parser} ? _parser ($opt{parser}) eq "csv"
		     : ($io_fil && $txt =~ m/\.(csv)$/i)) {
	$can{csv} or croak "CSV parser not installed";

	my $label = $io_fil ? $txt : "IO";

	$debug and print STDERR "Opening CSV $label\n";

	my @data = (
	    {	type	=> "csv",
		parser  => $can{csv},
		version	=> $can{csv}->VERSION,
		error	=> undef,
		quote   => '"',
		sepchar => ',',
		sheets	=> 1,
		sheet	=> { $label => 1 },
		},
	    {	label	=> $label,
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		merged  => [],
		},
	    );

	my ($sep, $quo, $in) = (",", '"');
	defined $opt{sep}   and $sep = $opt{sep};
	defined $opt{quote} and $quo = $opt{quote};
	if ($io_fil) {
	    unless (defined $opt{quote} && defined $opt{sep}) {
		open $in, "<", $txt or return;
		$_ = <$in>;

		$quo = defined $opt{quote} ? $opt{quote} : '"';
		$sep = # If explicitly set, use it
		   defined $opt{sep} ? $opt{sep} :
		       # otherwise start auto-detect with quoted strings
		       m/["0-9];["0-9;]/	? ";"  :
		       m/["0-9],["0-9,]/	? ","  :
		       m/["0-9]\t["0-9,]/	? "\t" :
		       # If neither, then for unquoted strings
		       m/\w;[\w;]/		? ";"  :
		       m/\w,[\w,]/		? ","  :
		       m/\w\t[\w,]/		? "\t" :
					      ","  ;
		close $in;
		}
	    open $in, "<", $txt or return;
	    }
	else {
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

    # From /etc/magic: Microsoft Office Document
    if ($io_txt && _parser ($opt{parser}) !~ m/^xlsx?$/ &&
		    $txt =~ m{^(\376\067\0\043
			       |\320\317\021\340\241\261\032\341
			       |\333\245-\0\0\0)}x) {
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
	$opt{parser} = "xls";
	}
    my $_parser;
    if ($opt{parser} ? ($_parser = _parser ($opt{parser})) =~ m/^xlsx?$/
		     : ($io_fil && $txt =~ m/\.(xlsx?)$/i && ($_parser = $1))) {
	my $parse_type = $_parser =~ m/x$/i ? "XLSX" : "XLS";
	my $parser  = $can{lc $parse_type} or
	    croak "Parser for $parse_type is not installed";
	$debug and print STDERR "Opening $parse_type \$txt\n";
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
	my @data = ( {
	    type	=> lc $parse_type,
	    parser	=> $can{lc $parse_type},
	    version	=> $can{lc $parse_type}->VERSION,
	    error	=> undef,
	    sheets	=> $oBook->{SheetCount} || 0,
	    sheet	=> {},
	    } );
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
	    : $can{xlsx} =~ m/::XLSX/
	    ? Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new
	    : Spreadsheet::ParseExcel::FmtDefault->new
	    : Spreadsheet::ParseExcel::FmtDefault->new;

	$debug and print STDERR "\t$data[0]{sheets} sheets\n";
	foreach my $oWkS (@{$oBook->{Worksheet}}) {
	    $opt{clip} and !defined $oWkS->{Cells} and next; # Skip empty sheets
	    my %sheet = (
		label	=> $oWkS->{Name},
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		merged  => [],
		);
	    defined $sheet{label}  or  $sheet{label}  = "-- unlabeled --";
	    exists $oWkS->{MaxRow} and $sheet{maxrow} = $oWkS->{MaxRow} + 1;
	    exists $oWkS->{MaxCol} and $sheet{maxcol} = $oWkS->{MaxCol} + 1;
	    $sheet{merged} = [
		map  {  $_->[0] }
		sort {  $a->[1] cmp $b->[1] }
		map  {[ $_, pack "NNNN", @$_          ]}
		map  {[ map { $_ + 1 } @{$_}[1,0,3,2] ]}
		@{$oWkS->get_merged_areas || []}];
	    my $sheet_idx = 1 + @data;
	    $debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} x $sheet{maxcol}\n";
	    if (exists $oWkS->{MinRow}) {
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
			defined (my $val = $oWkC->{Val})  or next;
			my $cell = cr2cell ($c + 1, $r + 1);
			$opt{rc}    and $sheet{cell}[$c + 1][$r + 1] = $val;	# Original

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
			    $sheet{$cell} = $FmT && exists $def_fmt{$FmT->{FmtIdx}}
				? $oFmt->ValFmt ($oWkC, $oBook)
				: $oWkC->Value;
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
				merged  => $oWkC->is_merged || 0,
				hidden  => $FmT->{Hidden}   || 0,
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
				};
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
	my @data = (
	    {	type	=> "sc",
		parser	=> "Spreadsheet::Read",
		version	=> $VERSION,
		error	=> undef,
		sheets	=> 1,
		sheet	=> { sheet => 1 },
		},
	    {	label	=> "sheet",
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		merged  => [],
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
	my $sxc_options = { %parser_opts, OrderBySheet => 1 }; # New interface 0.20 and up
	my $sxc;
	   if ($txt =~ m/\.(sxc|ods)$/i) {
	    $debug and print STDERR "Opening \U$1\E $txt\n";
	    $sxc = Spreadsheet::ReadSXC::read_sxc      ($txt, $sxc_options)	or  return;
	    }
	elsif ($txt =~ m/\.xml$/i) {
	    $debug and print STDERR "Opening XML $txt\n";
	    $sxc = Spreadsheet::ReadSXC::read_xml_file ($txt, $sxc_options)	or  return;
	    }
	# need to test on pattern to prevent stat warning
	# on filename with newline
	elsif ($txt !~ m/^<\?xml/i and -f $txt) {
	    $debug and print STDERR "Opening XML $txt\n";
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
		    label  => $sheet->{label},
		    maxrow => scalar @sheet,
		    maxcol => 0,
		    cell   => [],
		    attr   => [],
		    merged => [],
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

1;

__END__
=head1 DESCRIPTION

Spreadsheet::Read tries to transparently read *any* spreadsheet and
return its content in a universal manner independent of the parsing
module that does the actual spreadsheet scanning.

For OpenOffice and/or LibreOffice this module uses L<Spreadsheet::ReadSXC>

For Microsoft Excel this module uses L<Spreadsheet::ParseExcel>,
L<Spreadsheet::ParseXLSX>, or L<Spreadsheet::XLSX> (discouraged).

For CSV this module uses L<Text::CSV_XS> or L<Text::CSV_PP>.

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
        type    => "xls",
        parser  => "Spreadsheet::ParseExcel",
        version => 0.59,
	error	=> undef,
        },
      # Entry 1 is the first sheet
      { label   => "Sheet 1",
        maxrow  => 2,
        maxcol  => 4,
        cell    => [ undef,
          [ undef, 1 ],
          [ undef, undef, undef, undef, undef, "Nugget" ],
          ],
        attr    => [],
        merged  => [],
        A1      => 1,
        B5      => "Nugget",
        },
      # Entry 2 is the second sheet
      { label   => "Sheet 2",
        :
        :

To keep as close contact to spreadsheet users, row and column 1 have
index 1 too in the C<cell> element of the sheet hash, so cell "A1" is
the same as C<cell> [1, 1] (column first). To switch between the two,
there are two helper functions available: C<cell2cr ()> and C<cr2cell ()>.

The C<cell> hash entry contains unformatted data, while the hash entries
with the traditional labels contain the formatted values (if applicable).

The control hash (the first entry in the returned array ref), contains
some spreadsheet meta-data. The entry C<sheet> is there to be able to find
the sheets when accessing them by name:

  my %sheet2 = %{$book->[$book->[0]{sheet}{"Sheet 2"}]};

=head2 Functions

=over 2

=item my $book = ReadData ($source [, option => value [, ... ]]);

=item my $book = ReadData ("file.csv", sep => ',', quote => '"');

=item my $book = ReadData ("file.xls", dtfmt => "yyyy-mm-dd");

=item my $book = ReadData ("file.ods");

=item my $book = ReadData ("file.sxc");

=item my $book = ReadData ("content.xml");

=item my $book = ReadData ($content);

=item my $book = ReadData ($fh, parser => "xls");

Tries to convert the given file, string, or stream to the data
structure described above.

Processing Excel data from a stream or content is supported through
a File::Temp temporary file or IO::Scalar when available.

ReadSXC does preserve sheet order as of version 0.20.

Currently supported options are:

=over 2

=item parser

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

Control the generation of named cells ("A1" etc). Default is true.

=item rc

Control the generation of the {cell}[c][r] entries. Default is true.

=item attr

Control the generation of the {attr}[c][r] entries. Default is false.
See L<Cell Attributes> below.

=item clip

If set, C<ReadData ()> will remove all trailing rows and columns per
sheet that have no visual data. If a sheet has no data at all, the
sheet will be skipped entirely when this attribute is true.

This option is only valid if C<cells> is true. The default value is
true if C<cells> is true, and false otherwise.

=item strip

If set, C<ReadData ()> will remove trailing- and/or leading-whitespace
from every field.

  strip  leading  strailing
  -----  -------  ---------
    0      n/a      n/a
    1     strip     n/a
    2      n/a     strip
    3     strip    strip

=item sep

Set separator for CSV. Default is comma C<,>.

=item quote

Set quote character for CSV. Default is C<">.

=item dtfmt

Set the format for MS-Excel date fields that are set to use the default
date format. The default format in Excel is 'm-d-yy', which is both
not year 2000 safe, nor very useful. The default is now 'yyyy-mm-dd',
which is more ISO-like.

Note that date formatting in MS-Excel is not reliable at all, as it will
store/replace/change the date field separator in already stored formats
if you change your locale settings. So the above mentioned default can
be either "m-d-yy" OR "m/d/yy" depending on what that specific character
happened to be at the time the user saved the file.

=item debug

Enable some diagnostic messages to STDERR.

The value determines how much diagnostics are dumped (using Data::Peek).
A value of 9 and higher will dump the entire structure from the back-end
parser.

=back

All other attributes/options will be passed to the underlying parser if
that parser supports attributes.

=item my $cell = cr2cell (col, row)

C<cr2cell ()> converts a C<(column, row)> pair (1 based) to the
traditional cell notation:

  my $cell = cr2cell ( 4, 14); # $cell now "D14"
  my $cell = cr2cell (28,  4); # $cell now "AB4"

=item my ($col, $row) = cell2cr ($cell)

C<cell2cr ()> converts traditional cell notation to a C<(column, row)>
pair (1 based):

  my ($col, $row) = cell2cr ("D14"); # returns ( 4, 14)
  my ($col, $row) = cell2cr ("AB4"); # returns (28,  4)

=item my @row = row ($sheet, $row)

=item my @row = Spreadsheet::Read::row ($book->[1], 3)

Get full row of formatted values (like C<< $sheet->{A3} .. $sheet->{G3} >>)

Note that the indexes in the returned list are 0-based.

C<row ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=item my @row = cellrow ($sheet, $row)

=item my @row = Spreadsheet::Read::cellrow ($book->[1], 3)

Get full row of unformatted values (like C<< $sheet->{cell}[1][3] .. $sheet->{cell}[7][3] >>)

Note that the indexes in the returned list are 0-based.

C<cellrow ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=item my @rows = rows ($sheet)

=item my @rows = Spreadsheet::Read::rows ($book->[1])

Convert C<{cell}>'s C<[column][row]> to a C<[row][column]> list.

Note that the indexes in the returned list are 0-based, where the
index in the C<{cell}> entry is 1-based.

C<rows ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=item parses ($format)

=item Spreadsheet::Read::parses ("CSV")

C<parses ()> returns Spreadsheet::Read's capability to parse the
required format.

C<parses ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=item my $rs_version = Version ()

=item my $v = Spreadsheet::Read::Version ()

Returns the current version of Spreadsheet::Read.

C<Version ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=back

=head2 Using CSV

In case of CSV parsing, C<ReadData ()> will use the first line of the file
to auto-detect the separation character if the first argument is a file and
both C<sep> and C<quote> are not passed as attributes. Text::CSV_XS (or
Text::CSV_PP) is able to automatically detect and use C<\r> line endings.

CSV can parse streams too, but be sure to pass C<sep> and/or C<quote> if
these do not match the default C<,> and C<">.

When an error is found in the CSV, it is automatically reported (to STDERR).
The structure will store the error in C<< $ss->[0]{error} >> as anonymous
list returned by C<< $csv->error_diag >>. See Text::CSV_XS for documentation.

 my $ss = ReadData ("bad.csv");
 $ss->[0]{error} and say $ss->[0]{error}[1];

=head2 Cell Attributes

If the constructor was called with C<attr> having a true value, effort
is made to analyze and store field attributes like this:

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

This has now been partially implemented, mainly for Excel, as the other
parsers do not (yet) support all of that. YMMV.

=head3 Merged cells

Note that only [Spreadsheet::ReadSXC] documents the use of merged cells,
and not in a way useful for the spreadsheet consumer.

CSV does not support merged cells (though future implementations of CSV
for the web might).

The documentation of merged areas in [Spreadsheet::ParseExcel] and
[Spreadsheet::ParseXLSX] can be found in [Spreadsheet::ParseExcel::Worksheet]
and [Spreadsheet::ParseExcel::Cell].

None of basic [Spreadsheet::XLSX], [Spreadsheet::ParseExcel], and
[Spreadsheet::ParseXLSX] manual pages mention merged cells at all.

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

In this example, there is no way to see if B2 is merged to A2 or
to B1 without analyzing all surrounding cells. This could as well
mean A2:A3, B1:C1, B2: C2, as A2:A3, B1:B2, C1:C2, as A2:A3, B1:C2.
Use the C<merged> entry described above to find out what fields are
merged to what other fields.

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
        -n          Number lines (prefix with column number)
        -A          Show field attributes in ANSI escapes
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
        -n          Number lines (prefix with column number)
        -A          Show field attributes in ANSI escapes
     Grep options:
        -i          Ignore case
        -w          Match whole words only
        -h[#]       Show # header lines
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

=head2 C<ss2tk>

Show a spreadsheet in a perl/Tk spreadsheet widget

 usage: ss2tk [-w <width>] [X11 options] file.xls [<pattern>]
        -w <width> use <width> as default column width (4)

=head2 C<xls2csv>

Convert a spreadsheet to CSV. This is just a small wrapper over C<xlscat>.

 usage: xls2csv [ -o file.csv ] file.xls

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

=item Other spreadsheet formats

I consider adding any spreadsheet interface that offers a usable API.

=item Add an OO interface

Consider making the ref an object, though I currently don't see the big
advantage (yet). Maybe I'll make it so that it is a hybrid functional /
OO interface.

=back

=head1 SEE ALSO

=over 2

=item Text::CSV_XS, Text::CSV_PP

http://metacpan.org/release/Text-CSV_XS ,
http://metacpan.org/release/Text-CSV_PP , and
http://metacpan.org/release/Text-CSV .

Text::CSV is a wrapper over Text::CSV_XS (the fast XS version) and/or
Text::CSV_PP (the pure perl version)

=item Spreadsheet::ParseExcel

http://metacpan.org/release/Spreadsheet-ParseExcel

=item Spreadsheet::ParseXLSX

http://metacpan.org/release/Spreadsheet-ParseXLSX

=item Spreadsheet::XLSX

http://metacpan.org/release/Spreadsheet-XLSX

This module is considered dead and deprecated. It is buggy and unmaintained.
B<PLEASE> consider to use L</Spradsheet::ParseXLSX> instead.

=item Spreadsheet::ReadSXC

http://metacpan.org/release/Spreadsheet-ReadSXC

=item Spreadsheet::BasicRead

http://metacpan.org/release/Spreadsheet-BasicRead
for xlscat likewise functionality (Excel only)

=item Spreadsheet::ConvertAA

http://metacpan.org/release/Spreadsheet-ConvertAA
for an alternative set of cell2cr () / cr2cell () pair

=item Spreadsheet::Perl

http://metacpan.org/release/Spreadsheet-Perl
offers a Pure Perl implementation of a spreadsheet engine. Users that want
this format to be supported in Spreadsheet::Read are hereby motivated to
offer patches. It's not high on my TODO-list.

=item Spreadsheet::CSV

http://metacpan.org/release/Spreadsheet-CSV
offers the interesting approach of seeing all supported spreadsheet formats
as if it were CSV, mimicking the Text::CSV_XS interface.

=item xls2csv

http://metacpan.org/release/xls2csv offers an alternative for my C<xlscat -c>,
in the xls2csv tool, but this tool focuses on character encoding
transparency, and requires some other modules.

=back

=head1 AUTHOR

H.Merijn Brand, <h.m.brand@xs4all.nl>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2005-2015 H.Merijn Brand

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
