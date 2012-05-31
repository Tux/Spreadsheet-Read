#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 5;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "No M\$-Excel parser found";

my $xls;

{   local *STDERR;	# We want the debug activated, but not shown
    open   STDERR, ">", "/dev/null" or die "/dev/null: $!\n";
    $xls = ReadData ("files/misc.xlsx",
	# All defaults reversed
	rc	=> 0,
	cells	=> 0,
	attr	=> 1,
	clip	=> 1,
	debug	=> 5,
	);
    }
ok ($xls,				"Open with options");
is ($xls->[0]{sheets}, 3,		"Sheet Count");

{   local *STDERR;	# We want the debug activated, but not shown
    open   STDERR, ">", "/dev/null" or die "/dev/null: $!\n";
    $xls = ReadData ("files/misc.xlsx",
	# All defaults reversed, but undef
	rc	=> undef,
	cells	=> undef,
	attr	=> 1,
	clip	=> 1,
	debug	=> 5,
	);
    }
ok ($xls,				"Open with options");
is (0+@{ $xls->[1]{cell}[1]}, 0,	"undef works as option value for 'rc'");
ok (!exists $xls->[1]{A1},		"undef works as option value for 'cells'");

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
