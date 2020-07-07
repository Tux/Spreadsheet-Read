#!/usr/bin/perl

use strict;
use warnings;

BEGIN { $ENV{SPREADSHEET_READ_ODS} = "Spreadsheet::ParseODS"; }

my     $tests = 266;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("ods") or
    plan skip_all => "Cannot use $ENV{SPREADSHEET_READ_ODS}";

my $ods;
ok ($ods = ReadData ("files/attr.ods", attr => 1, debug => 99), "Excel Attributes testcase");

SKIP: {
    ok (my $clr = $ods->[$ods->[0]{sheet}{Colours}], "colors");

    defined $clr->{attr}[2][2]{fgcolor} or
	skip "$ods->[0]{parser} $ods->[0]{version} does not reliably support colors yet", 255;

    is ($clr->{cell}[1][1],		"auto",	"Auto");
    is ($clr->{attr}[1][1]{fgcolor}, undef,	"Unspecified font color");
    is ($clr->{attr}[1][1]{bgcolor}, undef,	"Unspecified fill color");

    my @clr = ( [],
	[ "auto",	undef     ],
	[ "red",	"#ff0000" ],
	[ "green",	"#008000" ],
	[ "blue",	"#0000ff" ],
	[ "white",	"#ffffff" ],
	[ "yellow",	"#ffff00" ],
	[ "lightgreen",	"#00ff00" ],
	[ "lightblue",	"#00ccff" ],
	[ "gray",	"#808080" ],
	);
    foreach my $col (1 .. $#clr) {
	my $bg = $clr[$col][1];
	is ($clr->{cell}[$col][1],		$clr[$col][0],	"Column $col header");
	foreach my $row (1 .. $#clr) {
	    my $fg = $clr[$row][1];
	    is ($clr->{cell}[1][$row],	$clr[$row][0],	"Row $row header");
	    is ($clr->{attr}[$col][$row]{fgcolor}, $fg,	"FG ($col, $row)");
	    is ($clr->{attr}[$col][$row]{bgcolor}, $bg,	"BG ($col, $row)");
	    }
	}
    }

ok ($ods = Spreadsheet::Read->new ("files/attr.ods", attr => 1), "Attributes OO");
is ($ods->[1]{attr}[3][3]{fgcolor},		"#008000", "C3 Forground color direct");
is ($ods->sheet (1)->attr (3, 3)->{fgcolor},	"#008000", "C3 Forground color OO rc   hash");
is ($ods->sheet (1)->attr ("C3")->{fgcolor},	"#008000", "C3 Forground color OO cell hash");
is ($ods->sheet (1)->attr (3, 3)->fgcolor,	"#008000", "C3 Forground color OO rc   method");
is ($ods->sheet (1)->attr ("C3")->fgcolor,	"#008000", "C3 Forground color OO cell method");

is ($ods->[1]{attr}[3][3]{bogus_attribute},	undef, "C3 bogus attribute direct");
is ($ods->sheet (1)->attr ("C3")->{bogus_attr},	undef, "C3 bogus attribute OO hash");
is ($ods->sheet (1)->attr ("C3")->bogus_attr,	undef, "C3 bogus attribute OO method");

#unless ($ENV{AUTOMATED_TESTING}) {
#    Test::NoWarnings::had_no_warnings ();
#    $tests++;
#    }
done_testing ($tests);
