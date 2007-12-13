#!/usr/bin/perl

# Use the sc parser because it's a builtin

use strict;
use warnings;

use Test::More;

use Spreadsheet::Read;
if (Spreadsheet::Read::parses ("sc")) {
    plan tests => 80;
    }
else {
    plan skip_all => "No SquirelCalc parser found";
    }

# Base attributes
foreach my $onoff (0, 1) {
    for (   [ ],
	    [ rc   => $onoff ],
	    [ cell => $onoff ],
	    [ rc   => 0, cell => $onoff ],
	    [ rc   => 1, cell => $onoff ],
	    [ clip => $onoff ],
	    [ cell => 0, clip => $onoff ],
	    [ cell => 1, clip => $onoff ],
	    [ attr => $onoff ],
	    [ cell => 0, attr => $onoff ],
	    ) {
	my $ref = ReadData ("files/test.sc", @$_);
	ok ($ref, "Open with options (@$_)");
	ok (ref $ref, "Valid ref");
	   $ref = ReadData ("files/test.sc", { @$_ });
	ok ($ref, "Open with options {@$_}");
	ok (ref $ref, "Valid ref");
	}
    }

# TODO: test and catch unsupported option.
#       Currently they are silently ignored
