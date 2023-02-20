#!perl

use strict;
use warnings;

my $tests = 257;
use Test::More;
require Test::NoWarnings;

use Spreadsheet::Read;
Spreadsheet::Read::parses ("gnumeric")
    or plan skip_all => "No Gnumeric parser found";

my $book;
ok ($book = ReadData ("files/attr.gnumeric", attr => 1),
    "Gnumeric attributes testcase");

ok (my $clr = $book->[$book->[0]{sheet}{Colours}], "have the 'Colours' sheet");

is ($clr->{cell}[1][1],          "auto",    "Auto");
is ($clr->{attr}[1][1]{fgcolor}, '#000000', "Unspecified font color");
is ($clr->{attr}[1][1]{bgcolor}, '#FFFFFF', "Unspecified fill color");

my @clr = (
    [],
    ["auto",       undef],
    ["red",        "#FF0000"],
    ["green",      "#008000"],
    ["blue",       "#0000FF"],
    ["white",      "#FFFFFF"],
    ["yellow",     "#FFFF00"],
    ["lightgreen", "#00FF00"],
    ["lightblue",  "#00CCFF"],
    ["gray",       "#808080"],
    );

foreach my $col (1 .. $#clr) {
    my $bg = $clr[$col][1] || '#FFFFFF';
    is ($clr->{cell}[$col][1], $clr[$col][0], "Column $col header");
    foreach my $row (1 .. $#clr) {
	my $fg = $clr[$row][1] || '#000000';
	is ($clr->{cell}[1][$row], $clr[$row][0],   "Row $row header");
	is ($clr->{attr}[$col][$row]{fgcolor}, $fg, "FG ($col, $row)");
	is ($clr->{attr}[$col][$row]{bgcolor}, $bg, "BG ($col, $row)");
	}
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
