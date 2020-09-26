#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_ODS} = "Spreadsheet::ParseODS"; }

my     $tests = 128;
use     Test::More;
#require Test::NoWarnings;

use     Spreadsheet::Read;

my $parser = Spreadsheet::Read::parses ("ods") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_ODS}";

my $pv = $parser->VERSION;
print STDERR "# Parser: $parser-$pv\n";

my $notyet = $pv lt "0.25" and $tests -= 10;

{   my $ref;
    $ref = ReadData ("no_such_file.ods");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadData ("files/empty.ods", debug => 8);
    use Data::Dumper;
    #local $TODO = "Handling of empty files needs more definition";
    ok (!defined $ref, "Empty file")
        or diag Dumper $ref;
    }

my $content;
{   local $/;
    open my $ods, "<", "files/test.ods" or die "files/test.ods: $!";
    binmode $ods;
    $content = <$ods>;
    close   $ods;
    }

my $ods;
foreach my $base ( [ "files/test.ods",	"Read/Parse ods file"	],
#		   [ $content,		"Parse ods data"	],
		   ) {
    my ($txt, $msg) = @$base;
    ok ($ods = ReadData ($txt),	$msg);

    ok (1, "Base values");
    is (ref $ods,		"ARRAY",	"Return type");
    is ($ods->[0]{type},	"ods",		"Spreadsheet type");
    is ($ods->[0]{sheets},	2,		"Sheet count");
    is (ref $ods->[0]{sheet},	"HASH",		"Sheet list");
    if(! is (scalar keys %{$ods->[0]{sheet}},
				2,		"Sheet list count")) {
        diag $_ for keys %{$ods->[0]{sheet}};
        use Data::Dumper;
        diag Dumper $ods;
    };
    my $xvsn = $ods->[0]{version};
    $xvsn =~ s/_[0-9]+$//; # remove beta part
    cmp_ok ($xvsn, ">=",	0.07,		"Parser version");

    ok (1, "Defined fields");
    foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
	my ($c, $r) = cell2cr ($cell);
    use Data::Dumper;
	is ($ods->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell")
        or diag Dumper $ods->[1];
	is ($ods->[1]{$cell},		$cell,	"Formatted   cell $cell");
	}

    ok (1, "Undefined fields");
    foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($ods->[1]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($ods->[1]{$cell},		undef,	"Formatted   cell $cell");
	}
    my @row = Spreadsheet::Read::rows ($ods->[1]);
    is (scalar @row,		4,	"Row'ed rows");
    is (scalar @{$row[3]},	4,	"Row'ed columns");
    is ($row[0][3],		"D1",	"Row'ed value D1");
    is ($row[3][2],		"C4",	"Row'ed value C4");
    }

# Tests for empty thingies
ok ($ods = ReadData ("files/values.ods"), "True/False values");
ok (my $ss = $ods->[1], "first sheet");
is ($ss->{cell}[1][1],	"A1",  "unformatted plain text");
is ($ss->{cell}[2][1],	" ",   "unformatted space") unless $notyet;
is ($ss->{cell}[3][1],	undef, "unformatted empty");
is ($ss->{cell}[4][1],	"0",   "unformatted numeric 0")
    or diag Dumper $ss->{cell}->[4]->[1];
is ($ss->{cell}[5][1],	"1",   "unformatted numeric 1");
is ($ss->{cell}[6][1],	"'",   "unformatted a single '");
is ($ss->{A1},		"A1",  "formatted plain text");
is ($ss->{B1},		" ",   "formatted space")	unless $notyet;
is ($ss->{C1},		undef, "formatted empty");
is ($ss->{D1},		"0",   "formatted numeric 0");
is ($ss->{E1},		"1",   "formatted numeric 1");
is ($ss->{F1},		"'",   "formatted a single '");

{   # RT#74976] Error Received when reading empty sheets
    foreach my $strip (0 .. 3) {
	my $ref = ReadData ("files/blank.ods", strip => $strip);
	ok ($ref, "File with no content - strip $strip");
	}
    }
   #use DP;die DDumper $ref;

# blank.ods has only one sheet with A1 filled with ' '
{   my  $ref = ReadData ("files/blank.ods", clip => 0, strip => 0);
    ok ($ref, "!clip strip 0");
    is ($ref->[1]{maxrow},     1,     "maxrow 1");
    is ($ref->[1]{maxcol},     1,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], " ",   "(1, 1) = ' '")	unless $notyet;
    is ($ref->[1]{A1},         " ",   "A1     = ' '")	unless $notyet;
        $ref = ReadData ("files/blank.ods", clip => 0, strip => 1);
    ok ($ref, "!clip strip 1");
    is ($ref->[1]{maxrow},     1,     "maxrow 1");
    is ($ref->[1]{maxcol},     1,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], "",    "blank (1, 1)")	unless $notyet;
    is ($ref->[1]{A1},         "",    "undef A1")	unless $notyet;
	$ref = ReadData ("files/blank.ods", clip => 0, strip => 1, cells => 0);
    ok ($ref, "!clip strip 1");
    is ($ref->[1]{maxrow},     1,     "maxrow 1");
    is ($ref->[1]{maxcol},     1,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], "",    "blank (1, 1)")	unless $notyet;
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = ReadData ("files/blank.ods", clip => 0, strip => 2,             rc => 0);
    ok ($ref, "!clip strip 2");
    is ($ref->[1]{maxrow},     1,     "maxrow 1");
    is ($ref->[1]{maxcol},     1,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         "",    "blank A1")	unless $notyet;
	$ref = ReadData ("files/blank.ods", clip => 0, strip => 3, cells => 0, rc => 0);
    ok ($ref, "!clip strip 3");
    is ($ref->[1]{maxrow},     1,     "maxrow 1");
    is ($ref->[1]{maxcol},     1,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");

	$ref = ReadData ("files/blank.ods", clip => 1, strip => 0);
    ok ($ref, " clip strip 0");
    my ($_mxr, $_mxc) = $pv lt "0.25" ? (0, 0) : (1, 1);
    is ($ref->[1]{maxrow},     $_mxr, "maxrow 1")
        or do { use Data::Dumper; diag Dumper $ref };
    is ($ref->[1]{maxcol},     $_mxc, "maxcol 1");
    is ($ref->[1]{cell}[1][1], " ",   "(1, 1) = ' '")	unless $notyet;
    is ($ref->[1]{A1},         " ",   "A1     = ' '")	unless $notyet;

	$ref = ReadData ("files/blank.ods", clip => 1, strip => 1);
    ok ($ref, " clip strip 1");
    is ($ref->[1]{maxrow},     0,     "maxrow 1");
    is ($ref->[1]{maxcol},     0,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], undef, "empty (1, 1)");
    is ($ref->[1]{A1},         undef, "empty A1");
	$ref = ReadData ("files/blank.ods", clip => 1, strip => 1, cells => 0);
    ok ($ref, " clip strip 1");
    is ($ref->[1]{maxrow},     0,     "maxrow 1");
    is ($ref->[1]{maxcol},     0,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], undef, "empty (1, 1)");
    is ($ref->[1]{A1},         undef, "empty A1");
	$ref = ReadData ("files/blank.ods", clip => 1, strip => 2,             rc => 0);
    ok ($ref, " clip strip 2");
    is ($ref->[1]{maxrow},     0,     "maxrow 1");
    is ($ref->[1]{maxcol},     0,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], undef, "empty (1, 1)");
    is ($ref->[1]{A1},         undef, "empty A1");
	$ref = ReadData ("files/blank.ods", clip => 1, strip => 3, cells => 0, rc => 0);
    ok ($ref, " clip strip 3");
    is ($ref->[1]{maxrow},     0,     "maxrow 1");
    is ($ref->[1]{maxcol},     0,     "maxcol 1");
    is ($ref->[1]{cell}[1][1], undef, "empty (1, 1)");
    is ($ref->[1]{A1},         undef, "empty A1");
    }

{   sub chk_test {
	my ($msg, $ods) = @_;

	is (ref $ods,		"ARRAY",	"Return type for $msg");
	is ($ods->[0]{type},	"ods",		"Spreadsheet type ODS");
	is ($ods->[0]{sheets},	2,		"Sheet count")
	} # chk_test

    my $data = $content;
    open my $fh, "<", "files/test.ods";
    binmode $fh;
    chk_test ( " FH    parser",    ReadData ( $fh,   parser => "ods")); close $fh;
    chk_test ("\\DATA  parser",    ReadData (\$data, parser => "ods"));
    chk_test ( " DATA  no parser", ReadData ( $data                  ));
    chk_test ( " DATA  parser",    ReadData ( $data, parser => "ods"));
    }

#unless ($ENV{AUTOMATED_TESTING}) {
#    Test::NoWarnings::had_no_warnings ();
#    $tests++;
#    }
done_testing ($tests);
