#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_CSV} = "Text::CSV_XS"; }

my     $tests = 133;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
my $parser = Spreadsheet::Read::parses ("csv") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_CSV}";

print STDERR "# Parser: $parser-", $parser->VERSION, "\n";

{   my $ref;
    $ref = ReadData ("no_such_file.csv");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadData ("files/empty.csv");
    ok (!defined $ref, "Empty file");
    }

my $csv;
ok ($csv = ReadData ("files/test.csv"),	"Read/Parse csv file");

ok (1, "Base values");
is (ref $csv,			"ARRAY",	"Return type");
is ($csv->[0]{type},		"csv",		"Spreadsheet type");
is ($csv->[0]{sheets},		1,		"Sheet count");
is (ref $csv->[0]{sheet},	"HASH",		"Sheet list");
is (scalar keys %{$csv->[0]{sheet}},
				1,		"Sheet list count");
cmp_ok ($csv->[0]{version},	">=",	0.01,	"Parser version");

is ($csv->[1]{maxrow},		5,		"Last row");
is ($csv->[1]{maxcol},		19,		"Last column");
is ($csv->[1]{cell}[$csv->[1]{maxcol}][$csv->[1]{maxrow}],
				"LASTFIELD",	"Last field");

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

ok ($csv = ReadData ("files/test_m.csv"),	"Read/Parse M\$ csv file");

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

ok ($csv = ReadData ("files/test_x.csv", sep => "=", quote => "_"),
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
	my $ref = ReadData ("files/blank.csv", strip => $strip);
	ok ($ref, "File with no content - strip $strip");
	}
    }

{   # RT#105197 - Strip wrong selection
    my  $ref = ReadData ("files/blank.csv", strip => 1);
    ok ($ref, "strip cells 1 rc 1");
    is ($ref->[1]{cell}[1][1], "",    "blank (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = ReadData ("files/blank.csv", strip => 1, cells => 0);
    ok ($ref, "strip cells 0 rc 1");
    is ($ref->[1]{cell}[1][1], "",    "blank (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = ReadData ("files/blank.csv", strip => 1,             rc => 0);
    ok ($ref, "strip cells 1 rc 0");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
	$ref = ReadData ("files/blank.csv", strip => 1, cells => 0, rc => 0);
    ok ($ref, "strip cells 0 rc 0");
    is ($ref->[1]{cell}[1][1], undef, "undef (1, 1)");
    is ($ref->[1]{A1},         undef, "undef A1");
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
