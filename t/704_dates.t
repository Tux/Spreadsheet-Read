#!/usr/bin/perl

use strict;
use warnings;

my $tests = 18;
use Test::More;
require Test::NoWarnings;

use Spreadsheet::Read;
Spreadsheet::Read::parses ("gnumeric")
    or plan skip_all => "Cannot use Spreadsheet::ReadGnumeric";

BEGIN { delete @ENV{qw( LANG LC_ALL LC_DATE )}; }

my $gnumeric;
ok ($gnumeric = ReadData ("files/Dates.gnumeric"), "Gnumeric date tests");

{
    ok (my $ss = $gnumeric->[1], "have sheet");

    my @date = (undef, 39668, 39672, 39790, 39673);
    foreach my $r (1 .. 4) {
	for (1 .. 4) {
	    is ($ss->{cell}[$_][$r], $date[$r], "Date value  row $r col $_");
	    }
	}
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
