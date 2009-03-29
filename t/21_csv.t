#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 12;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
	Spreadsheet::Read::parses ("csv") or
	    plan skip_all => "No CSV parser found";

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

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
