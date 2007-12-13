#!/usr/bin/perl

use strict;
use warnings;

use Test::More;

use Spreadsheet::Read;
if (Spreadsheet::Read::parses ("xls")) {
    plan "no_plan";
    }
else {
    plan skip_all => "No M\$-Excel parser found";
    }

my $xls;

{   local *STDERR;	# We want the debug activated, but not shown
    open   STDERR, ">", "/dev/null" or die "/dev/null: $!\n";
    $xls = ReadData ("files/misc.xls",
	# All defaults reversed
	rc		=> 0,
	cells	=> 0,
	attr	=> 1,
	clip	=> 1,
	debug	=> 5,
	);
    }
ok ($xls,					"Open with options");
is ($xls->[0]{sheets}, 3,			"Sheet Count");
