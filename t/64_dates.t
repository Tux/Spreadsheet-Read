#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 103;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "No M\$-Excel parser found";

BEGIN { delete @ENV{qw( LANG LC_ALL LC_DATE )}; }

my $xls;
ok ($xls = ReadData ("files/Dates.xlsx",
    attr => 1, dtfmt => "yyyy-mm-dd"), "Excel Date testcase");

my %fmt = (
    A1 => [ "8-Aug",			"d-mmm"			],
    A2 => [ "12-Aug",			"d-mmm"			],
    A3 => [ "8-Dec",			"d-mmm"			],
    A4 => [ "13-Aug",			"d-mmm"			],
    A6 => [ "Short: dd-MM-yyyy",	undef			],
    A7 => [ "2008-08-13",		"yyyy-mm-dd"		],
    B1 => [ 20080808,			"yyyymmdd"		],
    B2 => [ 20080812,			"yyyymmdd"		],
    B3 => [ 20081208,			"yyyymmdd"		],
    B4 => [ 20080813,			"yyyymmdd"		],
    B6 => [ "Long: ddd, dd MMM yyyy",	undef			],
    B7 => [ "Wed, 13 Aug 2008",		"ddd, dd mmm yyyy"	],
    C1 => [ "2008-08-08",		"yyyy-mm-dd"		],
    C2 => [ "2008-08-12",		"yyyy-mm-dd"		],
    C3 => [ "2008-12-08",		"yyyy-mm-dd"		],
    C4 => [ "2008-08-13",		"yyyy-mm-dd"		],
    C6 => [ "Default format 0x0E",	undef			],
    C7 => [ "8/13/08",			"m/d/yy"		],
    D1 => [ "08/08/2008",		"mm/dd/yyyy"		],
    D2 => [ "08/12/2008",		"mm/dd/yyyy"		],
    D3 => [ "12/08/2008",		"mm/dd/yyyy"		],
    D4 => [ "08/13/2008",		"mm/dd/yyyy"		],
    E1 => [ "08 Aug 2008",		undef			],
    E2 => [ "12 Aug 2008",		undef			],
    E3 => [ "08 Dec 2008",		undef			],
    E4 => [ "13 Aug 2008",		undef			],
    );

SKIP: {
    ok (my $ss   = $xls->[1],	"sheet");
    ok (my $attr = $ss->{attr},	"attr");

    defined $attr->[2][1]{format} or
	skip "$xls->[0]{parser} $xls->[0]{version} does not reliably support formats", 100;

    my @date = (undef, 39668,   39672,      39790,        39673);
    my @fmt  = (undef, "d-mmm", "yyyymmdd", "yyyy-mm-dd", "mm/dd/yyyy");
    foreach my $r (1 .. 4) {
	is ($ss->{cell}[$_][$r], $date[$r],    "Date value  row $r col $_") for 1 .. 4;

	is ($attr->[$_][$r]{type},   "date",   "Date type   row $r col $_")  for 1 .. 4;
	is ($attr->[$_][$r]{format}, $fmt[$_], "Date format row $r col $_")  for 1 .. 4;
	}

    foreach my $r (1..4,6..7) {
	foreach my $c (1..5) {
	    my $cell = cr2cell ($c, $r);
	    my $fmt  = $ss->{attr}[$c][$r]{format};
	    defined $ss->{$cell} or next;
	    is ($ss->{$cell}, $fmt{$cell}[0], "$cell content");
	    is ($fmt,         $fmt{$cell}[1], "$cell format");
	    }
	}
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
