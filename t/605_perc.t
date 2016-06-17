#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 77;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "No MS-Excel parser found";

my $xls;
ok ($xls = ReadData ("files/perc.xlsx", attr => 1), "Excel Percentage testcase");

my $ss   = $xls->[1];
my $attr = $ss->{attr};

foreach my $row (1 .. 19) {
    my @type = map { $ss->{attr}[$_][$row]{type} } 0 .. 3;
    is ($type[1], "numeric",	"Type A$row numeric");
    foreach my $col (2, 3) {	# Allow numeric for percentage in main test
	my $cell   = ("A".."C")[$col - 1].$row;
	my $expect = $type[$col] eq "numeric" ? "numeric" : "percentage";
	is ($type[$col], $expect, "Type B$row percentage");
	}

    SKIP: {
	$ss->{B18} =~ m/[.]/ and
	    skip "$xls->[0]{parser} $xls->[0]{version} has format problems", 1;
	my $i = int $ss->{"A$row"};
	# Allow edge case. rounding .5 will be different in -Duselongdouble perl
	my $f = $ss->{"B$row"};
	$row == 11 && $f eq "1%" and $i = 1;
	is ($f, "$i%",	"Formatted values for row $row\n");
	}
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
