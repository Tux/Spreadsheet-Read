#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 39;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
Spreadsheet::Read::parses ("xls") or
    plan skip_all => "No M\$-Excel parser found";

my $xls;
ok ($xls = ReadData ("files/attr.xls", attr => 1), "Excel Attributes testcase");

my $fmt = $xls->[$xls->[0]{sheet}{Format}];

is ($fmt->{B2},		"merged",	"Merged cell left    formatted");
is ($fmt->{C2},		"",		"Merged cell right   formatted");
is ($fmt->{cell}[2][2],	"merged",	"Merged cell left  unformatted");
is ($fmt->{cell}[3][2],	"",		"Merged cell right unformatted");
is ($fmt->{attr}[2][2]{merged}, 1,	"Merged cell left  merged");
is ($fmt->{attr}[3][2]{merged}, 1,	"Merged cell right merged");

is ($fmt->{B3},		"unlocked",	"Unlocked cell");
is ($fmt->{attr}[2][3]{locked}, 0,	"Unlocked cell not locked");
is ($fmt->{attr}[2][3]{merged}, 0,	"Unlocked cell not merged");
is ($fmt->{attr}[2][3]{hidden}, 0,	"Unlocked cell not hidden");

is ($fmt->{B4},		"hidden",	"Hidden cell");
is ($fmt->{attr}[2][4]{hidden}, 1,	"Hidden cell hidden");
is ($fmt->{attr}[2][4]{merged}, 0,	"Hidden cell not merged");

foreach my $r (1 .. 12) {
    is ($fmt->{cell}[1][$r], 12345,	"Unformatted valued A$r");
    }
is ($fmt->{attr}[1][1]{format}, undef,	"Default format");
is ($fmt->{cell}[1][1],  $fmt->{A1},	"Formatted valued A1");
is ($fmt->{cell}[1][10], $fmt->{A10},	"Formatted valued A10"); # String
foreach my $r (2 .. 9, 11, 12) {
    isnt ($fmt->{cell}[1][$r], $fmt->{"A$r"},	"Unformatted valued A$r");
    }
# Not yet. needs more digging
#foreach my $r (2 .. 12) {
#    ok (defined $fmt->{attr}[1][$r]{format},	"Defined format A$r");
#    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
