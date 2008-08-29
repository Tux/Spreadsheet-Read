#!/usr/bin/perl

use strict;
use warnings;

use Test::More;

use Spreadsheet::Read;
if (Spreadsheet::Read::parses ("xls")) {
    plan tests => 49;
    }
else {
    plan skip_all => "No M\$-Excel parser found";
    }

my $xls;
ok ($xls = ReadData ("files/Dates.xls", attr => 1), "Excel Date testcase");

SKIP: {
    my $ss   = $xls->[1];
    my $attr = $ss->{attr};

    my @date = (undef, 39668, 39672,      39790, 39673);
    my @fmt  = (undef, undef, "yyyymmdd", undef, "mm/dd/yyyy");
    foreach my $r (1 .. 4) {
	is ($ss->{cell}[$_][$r], $date[$r],	"Date value  row $r col $_") for 1 .. 4;

	SKIP: {
	    $xls->[0]{version} <= 0.31 and skip "Date types unreliable, please upgrade $xls->[0]{parser}", 8;
	    is ($attr->[$_][$r]{type},   "date",   "Date type   row $r col $_")  for 1 .. 4;
	    is ($attr->[$_][$r]{format}, $fmt[$_], "Date format row $r col $_")  for 1 .. 4;
	    }
	}
    }
