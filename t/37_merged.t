#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 22;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;

my $parser = Spreadsheet::Read::parses ("xls") or
    plan skip_all => "No M\$-Excel parser found";

ok (my $ss = ReadData ("files/merged.xls", attr => 1)->[1], "Read merged xls");

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

ok ($ss = ReadData ("files/merged.xls", attr => 1, merge => 1)->[1], "Read merged xlsx");

is ($ss->{attr}[1][1]{merged}, 0,    "unmerged A1");
is ($ss->{attr}[2][1]{merged}, "B1", "  merged B1");
is ($ss->{attr}[3][1]{merged}, "B1", "  merged C1");
is ($ss->{attr}[1][2]{merged}, "A2", "  merged A2");
is ($ss->{attr}[2][2]{merged}, "B1", "  merged B2");
is ($ss->{attr}[3][2]{merged}, "B1", "  merged C2");
is ($ss->{attr}[1][3]{merged}, "A2", "  merged A3");
is ($ss->{attr}[2][3]{merged}, 0,    "unmerged B3");
is ($ss->{attr}[3][3]{merged}, 0,    "unmerged C3");

is_deeply ($ss->{merged}, [[1,2,1,3],[2,1,3,2]], "Merged areas");

done_testing;
