#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_XLSX} = "Excel::ValueReader::XLSX"; }

my     $tests = 3;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_XLSX}";

ok (my $xls = ReadData ("files/perc.xlsx", attr => 1), "Excel Percentage testcase");
ok (my $ss  = $xls->[1], "We have a sheet");
ok (1, "No percentage support in $xls->[0]{parser}");

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
