#!/usr/bin/perl

use strict;
use warnings;

# OO version of 200_csv.t

my     $tests = 174;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
my $parser = Spreadsheet::Read::parses ("csv") or
    plan skip_all => "No CSV parser found";

{   my $ref;
    ok ($ref = Spreadsheet::Read->new ("no_such_file.csv"), "Open nonexisting file");
    is ($ref->[0]{sheets}, 0,	"No sheets read");
    ok ($ref = Spreadsheet::Read->new ("files/empty.csv"),  "Open empty file");
    is ($ref->[0]{sheets}, 0,	"No sheets read");
    }

my $csv;
ok ($csv = Spreadsheet::Read->new ("files/test.csv"),	"Read/Parse csv file");

ok (1, "Base values");
is (ref $csv,			"Spreadsheet::Read",	"Return type");
is ($csv->[0]{type},		"csv",			"Spreadsheet type");
is ($csv->[0]{sheets},		1,			"Sheet count");
is (ref $csv->[0]{sheet},	"HASH",			"Sheet list");
is (scalar keys %{$csv->[0]{sheet}}, 1,			"Sheet list count");
cmp_ok ($csv->[0]{version},	">=",	0.01,		"Parser version");

ok (my $sheet = $csv->sheet (1),			"Sheet 1");
is ($sheet->maxrow,		5,			"Last row");
is ($sheet->maxcol,		19,			"Last column");
is ($sheet->cell ($sheet->maxcol, $sheet->maxrow),
				"LASTFIELD",		"Last field");

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

{   # RT#105197 - Strip wrong selection
    my  $ref = Spreadsheet::Read->new ("files/blank.csv", strip => 1);
    ok ($ref, "strip cells 1 rc 1");
    is ($ref->[1]{cell}[1][1], "",    "blank (1, 1)");
    is ($ref->[1]{A1},         "",    "blank A1");
	$ref = Spreadsheet::Read->new ("files/blank.csv", strip => 1, cells => 0);
    ok ($ref, "strip cells 0 rc 1");
    is ($ref->[1]{cell}[1][1], "",    "blank (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = Spreadsheet::Read->new ("files/blank.csv", strip => 1,             rc => 0);
    ok ($ref, "strip cells 1 rc 0");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},          "",    "blank A1");
	$ref = Spreadsheet::Read->new ("files/blank.csv", strip => 1, cells => 0, rc => 0);
    ok ($ref, "strip cells 0 rc 0");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
    }

ok ($csv = Spreadsheet::Read->new ("files/test.csv"),	"Read/Parse csv file");
ok ($csv->add ("files/test.csv"), "Add the same file");
is ($csv->sheets, 2, "Two sheets");
is_deeply ([ $csv->sheets ], [qw( files/test.csv files/test.csv[2] )], "Sheet names");

is_deeply ($csv->sheet ("files/test.csv"), $csv->sheet (2), "Compare sheets");

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
