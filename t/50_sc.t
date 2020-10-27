#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 48;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
my $parser = Spreadsheet::Read::parses ("sc") or
    plan skip_all => "No SquirrelCalc parser found";

print STDERR "# Parser: $parser-", $parser->VERSION, "\n";

{   my $ref;
    $ref = ReadData ("no_such_file.sc");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadData ("files/empty.sc");
    ok (!defined $ref, "Empty file");
    }

my $content;
{   local $/;
    open my $sc, "<", "files/test.sc" or die "files/test.sc: $!\n";
    binmode $sc;
    $content = <$sc>;
    close   $sc;
    isnt ($content, undef, "Content is defined");
    isnt ($content, "",    "Content is filled");
    }

foreach my $txt ("files/test.sc", $content) {
    foreach my $clip (0, 2) {
	my $sc;
	ok ($sc = ReadData ($txt, clip => $clip, strip => 2),
	    "Read/Parse sc file ".($clip?"clipped":"unclipped"));

	ok (1, "Base values");
	is (ref $sc,			"ARRAY",	"Return type");
	is ($sc->[0]{type},		"sc",		"Spreadsheet type");
	is ($sc->[0]{sheets},		1,		"Sheet count");
	is (ref $sc->[0]{sheet},	"HASH",		"Sheet list");
	is (scalar keys %{$sc->[0]{sheet}},
					1,		"Sheet list count");
	is ($sc->[0]{version}, $Spreadsheet::Read::VERSION, "Parser version");

	is ($sc->[1]{maxcol},		10 - $clip,	"Columns");
	is ($sc->[1]{maxrow},		28 - $clip,	"Rows");
	is ($sc->[1]{cell}[1][22],	"  Workspace",	"Just checking one cell");
	}
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
