#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 71;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "No M\$-Excel parser found";

BEGIN { delete @ENV{qw( LANG LC_ALL LC_DATE )}; }

my $xls;
ok ($xls = ReadData ("files/Dates.xlsx", attr => 1, dtfmt => "yyyy-mm-dd"), "Excel Date testcase");

SKIP: {
    ok (my $ss   = $xls->[1],	"sheet");
    ok (my $attr = $ss->{attr},	"attr");

    defined $attr->[2][1]{format} or
	skip "$xls->[0]{parser} $xls->[0]{version} does not reliably support formats", 68;

    my @date = (undef, 39668, 39672,      39790,        39673);
    my @fmt  = (undef, undef, "yyyymmdd", "yyyy-mm-dd", "mm/dd/yyyy");
    foreach my $r (1 .. 4) {
	is ($ss->{cell}[$_][$r], $date[$r],	"Date value  row $r col $_") for 1 .. 4;

	is ($attr->[$_][$r]{type},   "date",   "Date type   row $r col $_")  for 1 .. 4;
	is ($attr->[$_][$r]{format}, $fmt[$_], "Date format row $r col $_")  for 1 .. 4;
	}

    is ($ss->{A1},	 "8-Aug",	"Cell content A1");
    is ($ss->{A2},	"12-Aug",	"Cell content A2");
    is ($ss->{A3},	 "8-Dec",	"Cell content A3");
    is ($ss->{A4},	"13-Aug",	"Cell content A4");

    is ($ss->{B1},	20080808,	"Cell content B1");
    is ($ss->{B2},	20080812,	"Cell content B2");
    is ($ss->{B3},	20081208,	"Cell content B3");
    is ($ss->{B4},	20080813,	"Cell content B4");

    is ($ss->{C1},	"2008-08-08",	"Cell content C1");
    is ($ss->{C2},	"2008-08-12",	"Cell content C2");
    is ($ss->{C3},	"2008-12-08",	"Cell content C3");
    is ($ss->{C4},	"2008-08-13",	"Cell content C4");

    is ($ss->{D1},	"08/08/2008",	"Cell content D1");
    is ($ss->{D2},	"08/12/2008",	"Cell content D2");
    is ($ss->{D3},	"12/08/2008",	"Cell content D3");
    is ($ss->{D4},	"08/13/2008",	"Cell content D4");

    is ($ss->{E1},	"08 Aug 2008",	"Cell content E1");
    is ($ss->{E2},	"12 Aug 2008",	"Cell content E2");
    is ($ss->{E3},	"08 Dec 2008",	"Cell content E3");
    is ($ss->{E4},	"13 Aug 2008",	"Cell content E4");
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
