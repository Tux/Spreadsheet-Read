#!/usr/bin/perl

use strict;
use warnings;

use Test::More;

use Spreadsheet::Read;
if (Spreadsheet::Read::parses ("xls")) {
    plan tests => 77;
    }
else {
    plan skip_all => "No M\$-Excel parser found";
    }

my $xls;
ok ($xls = ReadData ("files/perc.xls", attr => 1), "Excel Percentage testcase");

my $ss   = $xls->[1];
my $attr = $ss->{attr};

foreach my $row (1 .. 19) {
    is ($ss->{attr}[1][$row]{type}, "numeric",		"Type A$row numeric");
    is ($ss->{attr}[2][$row]{type}, "percentage",	"Type B$row percentage");
    is ($ss->{attr}[3][$row]{type}, "percentage",	"Type C$row percentage");

    my $i = int $ss->{"A$row"};
    is ("$i%", $ss->{"B$row"},		"Formatted values for row $row\n");
    }
