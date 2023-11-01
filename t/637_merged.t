#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_XLSX} = "Excel::ValueReader::XLSX"; }

my $tests = 2;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;

my $parser = Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_XLSX}";

# No merged cell support does not mean I should not be able to parse the file
ok (my $xls = ReadData ("files/merged.xlsx", attr => 1), "Read merged xlsx");
ok (1, "No merged cell support in $parser");

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
