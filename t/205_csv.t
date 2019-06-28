#!/usr/bin/perl

use strict;
use warnings;

# OO version of 200_csv.t

my     $tests = 261;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
my $parser = Spreadsheet::Read::parses ("csv") or
    plan skip_all => "No CSV parser found";

{   my $ref;
    is (Spreadsheet::Read->new ("no_such_file.csv"), undef, "Open nonexisting file");
    ok ($@, "No sheets read: $@"); # No such file or directory
    is (Spreadsheet::Read->new ("files/empty.csv"),  undef, "Open empty file");
    is ($@, "files/empty.csv is empty",	"No sheets read");
    }

my $csv;
ok ($csv = Spreadsheet::Read->new ("files/test.csv"),	"Read/Parse csv file");
ok ($csv->parses ("CSV"),		"Parses CSV");
is ($csv->cr2cell (1, 1),	"A1",	"method cr2cell");

ok (1, "Base values");
is (ref $csv,			"Spreadsheet::Read",	"Return type");
is ($csv->[0]{type},		"csv",			"Spreadsheet type");
is ($csv->[0]{sheets},		1,			"Sheet count");
is (ref $csv->[0]{sheet},	"HASH",			"Sheet list");
is (scalar keys %{$csv->[0]{sheet}}, 1,			"Sheet list count");
cmp_ok ($csv->[0]{version},	">=",	0.01,		"Parser version");

is_deeply ([$csv->cellrow (1, 1)], ["A1","B1","","D1",(undef) x 15], "row 1");
is ($csv->cellrow (1, 255), undef, "No such row   255");
is ($csv->cellrow (1, -55), undef, "No such row   -55");
is ($csv->cellrow (1,   0), undef, "No such row     0");
is ($csv->cellrow (0,   1), undef, "No such sheet   0");
is ($csv->cellrow (-2,  1), undef, "Wrong   sheet  -2");
is ($csv->cellrow (-20, 1), undef, "No such sheet -20");
is_deeply ([$csv->row     (1, 1)], ["A1","B1","","D1",(undef) x 15], "row 1");
is ($csv->row     (1, 255), undef, "No such row   255");
is ($csv->row     (1, -55), undef, "No such row   -55");
is ($csv->row     (1,   0), undef, "No such row     0");
is ($csv->row     (0,   1), undef, "No such sheet   0");
is ($csv->row     (-2,  1), undef, "Wrong   sheet  -2");
is ($csv->row     (-20, 1), undef, "No such sheet -20");

is ($csv->sheet (  0),	undef,				"Sheet   0");
is ($csv->sheet (255),	undef,				"Sheet 255");
is ($csv->sheet (-55),	undef,				"Sheet -55");
is ($csv->sheet ("="),	undef,				"Sheet '='");

is (Spreadsheet::Read::sheet (undef, 1), undef,	"Don't be silly");
is (Spreadsheet::Read::sheet ($csv, -9), undef,	"Don't be silly");

ok (my $sheet = $csv->sheet (1),			"Sheet 1");
is ($sheet->maxrow,		5,			"Last row");
is ($sheet->maxcol,		19,			"Last column");
is ($sheet->cell ($sheet->maxcol, $sheet->maxrow),
				"LASTFIELD",		"Last field");

is_deeply ([$sheet->cellrow (1)], ["A1","B1","","D1",(undef) x 15], "row 1");
is ($sheet->cellrow (255), undef, "No such row 255");
is ($sheet->cellrow (-55), undef, "No such row -55");
is ($sheet->cellrow (  0), undef, "No such row   0");
is_deeply ([$sheet->row     (1)], ["A1","B1","","D1",(undef) x 15], "row 1");
is ($sheet->row     (255), undef, "No such row 255");
is ($sheet->row     (-55), undef, "No such row -55");
is ($sheet->row     (  0), undef, "No such row   0");

is_deeply ([$sheet->cellcolumn (1)], ["A1","A2","A3","A4",""], "col 1");
is ($sheet->cellcolumn (255), undef, "No such col 255");
is ($sheet->cellcolumn (-55), undef, "No such col -55");
is ($sheet->cellcolumn (  0), undef, "No such col   0");
is_deeply ([$sheet->column     (1)], ["A1","A2","A3","A4",""], "col 1");
is ($sheet->column     (255), undef, "No such col 255");
is ($sheet->column     (-55), undef, "No such col -55");
is ($sheet->column     (  0), undef, "No such col   0");

ok (my @rows = $sheet->rows, "All rows");
is ($rows[0][0], "A1", "A1");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = $csv->cell2cr ($cell);
    is ($sheet->{cell}[$c][$r],	$cell,	"Unformatted cell $cell direct");
    is ($sheet->{$cell},	$cell,	"Formatted   cell $cell direct");
    is ($sheet->cell ($c, $r),	$cell,	"Unformatted cell $cell method");
    is ($sheet->cell ($cell),	$cell,	"Formatted   cell $cell method");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = $csv->cell2cr ($cell);
    is ($sheet->{cell}[$c][$r],	"",   	"Unformatted cell $cell direct");
    is ($sheet->{$cell},	"",   	"Formatted   cell $cell direct");
    is ($sheet->cell ($c, $r),	"",   	"Unformatted cell $cell method");
    is ($sheet->cell ($cell),	"",   	"Formatted   cell $cell method");
    }

ok ($csv = Spreadsheet::Read->new ("files/test_m.csv"),	"Read/Parse M\$ csv file");
ok ($sheet = $csv->sheet (1),				"Sheet 1");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = $sheet->cell2cr ($cell);
    is ($sheet->cell ($c, $r),	$cell,	"Unformatted cell $cell");
    is ($sheet->cell ($cell),	$cell,	"Formatted   cell $cell");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = $sheet->cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	"",   	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		"",   	"Formatted   cell $cell");
    }

ok ($csv = Spreadsheet::Read->new ("files/test_x.csv", sep => "=", quote => "_"),
					    "Read/Parse strange csv file");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		$cell,	"Formatted   cell $cell");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	"",   	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		"",   	"Formatted   cell $cell");
    }

{   # RT#74976 - Error Received when reading empty sheets
    foreach my $strip (0 .. 3) {
	my $ref = Spreadsheet::Read->new ("files/blank.csv", strip => $strip);
	ok ($ref, "File with no content - strip $strip");
	}
    }

# blank.csv has only one sheet with A1 filled with ' '
{   my  $ref = ReadData ("files/blank.csv", clip => 0, strip => 0);
    ok ($ref, "!clip strip 0");
    is ($ref->[1]{maxrow},     3,     "maxrow 3");
    is ($ref->[1]{maxcol},     4,     "maxcol 4");
    is ($ref->[1]{cell}[1][1], " ",   "(1, 1) = ' '");
    is ($ref->[1]{A1},         " ",   "A1     = ' '");
        $ref = ReadData ("files/blank.csv", clip => 0, strip => 1);
    ok ($ref, "!clip strip 1");
    is ($ref->[1]{maxrow},     3,     "maxrow 3");
    is ($ref->[1]{maxcol},     4,     "maxcol 4");
    is ($ref->[1]{cell}[1][1], "",    "blank (1, 1)");
    is ($ref->[1]{A1},         "",    "undef A1");
	$ref = ReadData ("files/blank.csv", clip => 0, strip => 1, cells => 0);
    ok ($ref, "!clip strip 1");
    is ($ref->[1]{maxrow},     3,     "maxrow 3");
    is ($ref->[1]{maxcol},     4,     "maxcol 4");
    is ($ref->[1]{cell}[1][1], "",    "blank (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = ReadData ("files/blank.csv", clip => 0, strip => 2,             rc => 0);
    ok ($ref, "!clip strip 2");
    is ($ref->[1]{maxrow},     3,     "maxrow 3");
    is ($ref->[1]{maxcol},     4,     "maxcol 4");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         "",    "blank A1");
	$ref = ReadData ("files/blank.csv", clip => 0, strip => 3, cells => 0, rc => 0);
    ok ($ref, "!clip strip 3");
    is ($ref->[1]{maxrow},     3,     "maxrow 3");
    is ($ref->[1]{maxcol},     4,     "maxcol 4");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");

	$ref = ReadData ("files/blank.csv", clip => 1, strip => 0);
    ok ($ref, " clip strip 0");
    is ($ref->[1]{maxrow},     1,     "maxrow 3");
    is ($ref->[1]{maxcol},     1,     "maxcol 4");
    is ($ref->[1]{cell}[1][1], " ",   "(1, 1) = ' '");
    is ($ref->[1]{A1},         " ",   "A1     = ' '");

	$ref = ReadData ("files/blank.csv", clip => 1, strip => 1);
    ok ($ref, " clip strip 1");
    is ($ref->[1]{maxrow},     0,     "maxrow 0");
    is ($ref->[1]{maxcol},     0,     "maxcol 0");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = ReadData ("files/blank.csv", clip => 1, strip => 1, cells => 0);
    ok ($ref, " clip strip 1");
    is ($ref->[1]{maxrow},     0,     "maxrow 0");
    is ($ref->[1]{maxcol},     0,     "maxcol 0");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = ReadData ("files/blank.csv", clip => 1, strip => 2,             rc => 0);
    ok ($ref, " clip strip 2");
    is ($ref->[1]{maxrow},     0,     "maxrow 0");
    is ($ref->[1]{maxcol},     0,     "maxcol 0");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = ReadData ("files/blank.csv", clip => 1, strip => 3, cells => 0, rc => 0);
    ok ($ref, " clip strip 3");
    is ($ref->[1]{maxrow},     0,     "maxrow 0");
    is ($ref->[1]{maxcol},     0,     "maxcol 0");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
    }

ok ($csv = Spreadsheet::Read->new ("files/test.csv"),	"Read/Parse csv file");
ok ($csv->add ("files/test.csv"), "Add the same file");
is ($csv->sheets, 2, "Two sheets");
is_deeply ([ $csv->sheets ], [qw( files/test.csv files/test.csv[2] )], "Sheet names");

is_deeply ($csv->sheet ("files/test.csv"), $csv->sheet (2), "Compare sheets");

ok (my $sheet2 = $csv->sheet (2),	"The new sheet");
is ($sheet2->label, "files/test.csv",	"Original label");
is ($sheet2->label ("Hello"), "Hello",	"New label");
ok (my $sheet3 = $csv->sheet ("Hello"),	"Found by new label");
is_deeply ($sheet2, $sheet3, "Compare sheets");

ok ($csv->add ("files/test.csv", label => "Test"), "Add with label");
is_deeply ([ $csv->sheets ], [qw( files/test.csv files/test.csv[2] Test )], "Sheet names");

is ($csv->col2label (4),             "D",  "col2label as book  method");
is ($csv->sheet (1)->col2label (27), "AA", "col2label as sheet method");

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
