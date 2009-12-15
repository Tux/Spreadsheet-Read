#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 118;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read qw( ReadData cell2cr row cellrow );
my $parser = Spreadsheet::Read::parses ("csv") or
    plan skip_all => "No CSV parser found";

sub ReadDataCSV
{
    ReadData (@_, parser => "csv");
    } # ReadDataCSV

{   my $ref;
    $ref = ReadDataCSV ("files/empty.txt");
    ok (!defined $ref, "Empty file");
    }

my $csv;
ok ($csv = ReadDataCSV ("files/test.txt"),	"Read/Parse csv file");

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

my @row = ("A3", "", "C3", "D3", (undef) x 15);
is_deeply ([ row     ($csv->[1], 3) ], \@row, "Formatted   row 3");
is_deeply ([ cellrow ($csv->[1], 3) ], \@row, "Unformatted row 3");

ok ($csv = ReadDataCSV ("files/test_m.txt"),	"Read/Parse M\$ csv file");

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

ok ($csv = ReadDataCSV ("files/test_x.txt", sep => "=", quote => "_"),
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

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
