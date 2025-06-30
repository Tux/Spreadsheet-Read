#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_XLSX} = "Excel::ValueReader::XLSX"; }

my     $tests = 4;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_XLSX}";

BEGIN { delete @ENV{qw( LANG LC_ALL LC_DATE )}; }

ok (my $xls = ReadData ("files/Dates.xlsx",
    attr => 1, dtfmt => "yyyy-mm-dd"), "Excel Date testcase");

ok (my $ss = $xls->[1],	    "sheet");
is_deeply ($ss->{attr}, [], "attr");

ok (1, "$xls->[0]{parser} $xls->[0]{version} does not support formats");

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
