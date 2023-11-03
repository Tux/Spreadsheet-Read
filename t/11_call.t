#!/usr/bin/perl

use strict;
use warnings;

use Test::More;
use Test::NoWarnings;

use Spreadsheet::Read;
    Spreadsheet::Read::parses ("sc") or
	plan skip_all => "No SquirrelCalc parser found";

plan tests => 243;

# Base attributes
foreach my $prs ([], [ parser => "sc" ], [ parser => "Spreadsheet::Read" ]) {
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
	my @attr = (@$_, @$prs);
	my $ref = ReadData ("files/test.sc",   @attr);
	ok ($ref, "Open with options ( @attr )");
	ok (ref $ref, "Valid ref");
	   $ref = ReadData ("files/test.sc", { @attr });
	ok ($ref, "Open with options { @attr }");
	ok (ref $ref, "Valid ref");
	}
    }
  }

{   my @err;
    local $SIG{__DIE__} = sub { push @err => @_ };
    my $p = eval {
	ReadData ("files/test.sc", parser => "Spreadsheet::Stupid");
	};
    is ($p, undef, "Cannot use unsupported parser");
    s/[\s\r\n]+at[\s\r\n]+\S+\s+line\s+\d+.*//s for @err;
    is_deeply (\@err,
	[ "I can open file files/test.sc, but I do not know how to parse it" ],
	"Cannot parse");
    }

# TODO: test and catch unsupported option.
#       Currently they are silently ignored
