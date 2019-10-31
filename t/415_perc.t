#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_ODS} = "Spreadsheet::ParseODS"; }

my     $tests = 77;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("ods") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_ODS}";

my $ods;
ok ($ods = ReadData ("files/perc.ods", attr => 1), "Excel Percentage testcase");

my $ss   = $ods->[1];
my $attr = $ss->{attr};

foreach my $row (1 .. 19) {
    my @type = map { $ss->{attr}[$_][$row]{type} } 0 .. 3;
    is ($ss->{attr}[1][$row]{type}, "numeric", "Type A$row numeric");
    foreach my $col (2, 3) {
	my $cell = ("A".."C")[$col - 1].$row;
	my $xvsn = $ods->[0]{version};
	$xvsn =~ s/_[0-9]+$//; # remove beta part
	my $expect = $xvsn < 0.23 && $type[$col] eq "numeric"
	    ? "numeric" : "percentage";
	is ($type[$col], $expect, "Type $cell percentage");
	}

    SKIP: {
	$ss->{B18} =~ m/[.]/ and
	    skip "$ods->[0]{parser} $ods->[0]{version} has format problems", 1;
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
