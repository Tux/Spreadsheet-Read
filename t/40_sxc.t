#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 301;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
my $parser = Spreadsheet::Read::parses ("sxc") or
    plan skip_all => "No SXC parser found";

print STDERR "# Parser: $parser-", $parser->VERSION, "\n";

my $content;
{   local $/;
    open my $xml, "<", "files/content.xml" or die "files/content.xml: $!\n";
    binmode $xml;
    $content = <$xml>;
    close   $xml;
    }

{   my $ref;
    $ref = ReadData ("no_such_file.sxc");
    ok (!defined $ref, "Nonexistent file");
    # Too noisy
    #eval { $ref = ReadData ("files/empty.sxc") };
    #ok (!defined $ref, "Empty file");
    #like ($@, qr/too short/);
    }

foreach my $base ( [ "files/test.sxc",		"Read/Parse sxc file" ],
		   [ "files/content.xml",	"Read/Parse xml file" ],
		   [ $content,			"Parse xml data" ],
		   ) {
    my ($txt, $msg) = @$base;
    my $sxc;
    ok ($sxc = ReadData ($txt), $msg);

    ok (1, "Base values");
    is (ref $sxc,		"ARRAY",	"Return type");
    is ($sxc->[0]{type},	"sxc",		"Spreadsheet type");
    is ($sxc->[0]{sheets},	2,		"Sheet count");
    is (ref $sxc->[0]{sheet},	"HASH",		"Sheet list");
    is (scalar keys %{$sxc->[0]{sheet}},
				2,		"Sheet list count");
    # This should match the version required in Makefile.PL's PREREQ_PM
    cmp_ok ($sxc->[0]{version}, ">=",	0.12,	"Parser version");

    ok (1, "Sheet 1");
    # Simple sheet with cells filled with the cell label:
    # -- -- -- --
    # A1 B1    D1
    # A2 B2
    # A3    C3 D3
    # A4 B4 C4

    ok (1, "Defined fields");
    foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
	is ($sxc->[1]{$cell},		$cell,	"Formatted   cell $cell");
	}

    ok (1, "Undefined fields");
    foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[1]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($sxc->[1]{$cell},		undef,	"Formatted   cell $cell");
	}

    ok (1, "Nonexistent fields");
    foreach my $cell (qw( A9 X6 B17 AB4 BE33 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[1]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($sxc->[1]{$cell},		undef,	"Formatted   cell $cell");
	}

    ok (1, "Sheet 2");
    # Sheet with merged cells and notes/annotations
    # x   x   x
    #   x   x 
    # x   x   x

    ok (1, "Defined fields");
    foreach my $cell (qw( A1 C1 E1 B2 D2 A3 C3 E3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[2]{cell}[$c][$r],	"x",	"Unformatted cell $cell");
	is ($sxc->[2]{$cell},		"x",	"Formatted   cell $cell");
	}

    ok (1, "Undefined fields");
    foreach my $cell (qw( B1 D1 A2 C2 E2 B3 D3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[2]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($sxc->[2]{$cell},		undef,	"Formatted   cell $cell");
	}

    ok (1, "Nonexistent fields");
    foreach my $cell (qw( A9 X6 B17 AB4 BE33 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[2]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($sxc->[2]{$cell},		undef,	"Formatted   cell $cell");
	}

    # Sheet order
    ok (exists $sxc->[0]{sheet}{Sheet1}, "Sheet labels in metadata");
    my @sheets = map { $sxc->[$_]{label} } 1 .. $sxc->[0]{sheets};
    SKIP: {
	$sxc->[0]{version} < 0.20 and
	    skip "Not supported", 1;
	is ("@sheets", "@{['Sheet1','Second Sheet']}", "Sheet order");
	}
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
