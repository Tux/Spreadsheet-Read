#!/usr/bin/perl

use strict;
use warnings;

use Test::More;

use Spreadsheet::Read;
if (Spreadsheet::Read::parses ("csv")) {
    plan "no_plan";
    }
else {
    plan skip_all => "No CSV parser found";
    }

my $csv;
ok ($csv = ReadData ("files/macosx.csv"),	"Read/Parse csv file");

#use DP; DDumper $csv;

is ($csv->[1]{maxrow},		16,		"Last row");
is ($csv->[1]{maxcol},		15,		"Last column");
is ($csv->[1]{cell}[$csv->[1]{maxcol}][$csv->[1]{maxrow}],
				"",		"Last field");

__END__
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
is ($csv->[0]{sepchar},	",",			"{sepchar}");
is ($csv->[0]{quote},	'"',			"{quote}");
is ($csv->[1]{C3},      "C3",			"cell C3");
