#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_XLSX} = "Spreadsheet::ParseXLSX"; }

my     $tests = 77;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_XLSX}";

my $xls;
ok ($xls = ReadData ("files/perc.xlsx", attr => 1), "Excel Percentage testcase");

my $ss   = $xls->[1];
my $attr = $ss->{attr};

foreach my $row (1 .. 19) {
    my @type = map { $ss->{attr}[$_][$row]{type} } 0 .. 3;
    is ($ss->{attr}[1][$row]{type}, "numeric",		"Type A$row numeric");
    if ($xls->[0]{version} < 0.23 && $type[2] eq "numeric") {
	is ($type[2], "numeric",	"Type B$row percentage");
	is ($type[3], "numeric",	"Type C$row percentage");
	}
    else {
	is ($type[2], "percentage",	"Type B$row percentage");
	is ($type[3], "percentage",	"Type C$row percentage");
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
