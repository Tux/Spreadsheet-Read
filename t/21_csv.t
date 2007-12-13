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
ok ($csv = ReadData ("files/test.csv"),		"Read/Parse csv file");
is ($csv->[0]{sepchar},	",",			"{sepchar}");
is ($csv->[0]{quote},	'"',			"{quote}");
is ($csv->[1]{C3},      "C3",			"cell C3");

ok ($csv = ReadData ("files/test_m.csv"),	"Read/Parse csv file (M$)");
is ($csv->[0]{sepchar},	";",			"{sepchar}");
is ($csv->[0]{quote},	'"',			"{quote}");
is ($csv->[1]{C3},      "C3",			"cell C3");

ok ($csv = ReadData ("files/test_t.csv", quote => "'"),
						"Read/Parse csv file (tabs)");
is ($csv->[0]{sepchar},	"\t",			"{sepchar}");
is ($csv->[0]{quote},	"'",			"{quote}");
is ($csv->[1]{C3},      "C3",			"cell C3");
