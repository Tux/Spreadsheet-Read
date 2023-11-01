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

my $xls;
ok ($xls = ReadData ("files/attr.xlsx", attr => 1), "Excel Attributes testcase");

my $parser = $xls->[0]{parser};

ok (my $fmt = $xls->[$xls->[0]{sheet}{Format}],	"format");
is ($fmt->{attr}[2][2]{merged}, undef, "$xls->[0]{parser} does not support attributes");

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
