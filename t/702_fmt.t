#!/usr/bin/perl

use strict;
use warnings;

my $tests = 30;
use Test::More;
require Test::NoWarnings;

use Spreadsheet::Read;
Spreadsheet::Read::parses ("gnumeric")
    or plan skip_all => "No Gnumeric parser found";

my $book;
ok ($book = ReadData ("files/attr.gnumeric", attr => 1),
    "Gnumeric attributes testcase");

{   ok (my $fmt = $book->[$book->[0]{sheet}{Format}], "have 'Format' sheet");

    is ($fmt->{B2},                 "merged", "Merged cell left    formatted");
    is ($fmt->{C2},                 undef,    "Merged cell right   formatted");
    is ($fmt->{cell}[2][2],         "merged", "Merged cell left  unformatted");
    is ($fmt->{cell}[3][2],         undef,    "Merged cell right unformatted");
    is ($fmt->{attr}[2][2]{merged}, 1,        "Merged cell left  merged");
    is ($fmt->{attr}[3][2]{merged}, 1,        "Merged cell right merged");

    is ($fmt->{B3},                 "unlocked", "Unlocked cell");
    is ($fmt->{attr}[2][3]{locked}, 0,          "Unlocked cell not locked");
    is ($fmt->{attr}[2][3]{merged}, undef,      "Unlocked cell not merged");
    is ($fmt->{attr}[2][3]{hidden}, 1,          "Unlocked cell is hidden");

    is ($fmt->{B4},                 "hidden", "Hidden cell");
    is ($fmt->{attr}[2][4]{hidden}, 1,        "Hidden cell hidden");
    is ($fmt->{attr}[2][4]{merged}, undef,    "Hidden cell not merged");

    foreach my $r (1 .. 12) {
	is ($fmt->{cell}[1][$r], 12345, "Unformatted valued A$r");
	}
    is ($fmt->{attr}[1][1]{format}, "General",  "Default format");
    is ($fmt->{cell}[1][1],         $fmt->{A1}, "Formatted valued A1");
    is ($fmt->{cell}[1][10], $fmt->{A10}, "Formatted valued A10");    # String
    # There's no point in testing the formatted values here, because Gnumeric
    # gives them to us unformatted only.
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
