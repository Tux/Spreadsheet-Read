#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_XLSX} = "Spreadsheet::XLSX::Reader::LibXML"; }

my     $tests = 11;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;

my $parser = Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_XLSX}";
my $pv = $parser->VERSION;

$pv le "v0.38.12" and
    plan skip_all => "No reliable merged cell support in version $parser";

ok (my $ss = ReadData ("files/merged.xlsx", attr => 1)->[1], "Read merged xlsx");

is ($ss->{attr}[1][1]{merged}, 0, "unmerged A1");
is ($ss->{attr}[2][1]{merged}, 1, "  merged B1");
is ($ss->{attr}[3][1]{merged}, 1, "  merged C1");
is ($ss->{attr}[1][2]{merged}, 1, "  merged A2");
is ($ss->{attr}[2][2]{merged}, 1, "  merged B2");
is ($ss->{attr}[3][2]{merged}, 1, "  merged C2");
is ($ss->{attr}[1][3]{merged}, 1, "  merged A3");
is ($ss->{attr}[2][3]{merged}, 0, "unmerged B3");
is ($ss->{attr}[3][3]{merged}, 0, "unmerged C3");

is_deeply ($ss->{merged}, [[1,2,1,3],[2,1,3,2]], "Merged areas");

done_testing;
