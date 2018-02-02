#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_XLSX} = "Test::More"; }

my     $tests = 2;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;

is (Spreadsheet::Read::parses ("xlsx"), 0, "Invalid module name for xlsx");
like ($@, qr/^Test::More is not supported/, "Error reason");

done_testing ($tests);
