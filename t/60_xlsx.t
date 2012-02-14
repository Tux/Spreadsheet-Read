#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 66;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;

my $parser = Spreadsheet::Read::parses ("xlsx") or
    plan skip_all => "No M\$-Excel parser found";

print STDERR "# Parser: $parser-", $parser->VERSION, "\n";

{   my $ref;
    $ref = ReadData ("no_such_file.xlsx");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadData ("files/empty.xlsx");
    ok (!defined $ref, "Empty file");
    }

my $content;
{   local $/;
    open my $xls, "<", "files/test.xlsx" or die "files/test.xlsx: $!";
    binmode $xls;
    $content = <$xls>;
    close   $xls;
    }

my $xls;
foreach my $base ( [ "files/test.xlsx",	"Read/Parse xlsx file"	],
#		   [ $content,		"Parse xlsx data"	],
		   ) {
    my ($txt, $msg) = @$base;
    ok ($xls = ReadData ($txt),	$msg);

    ok (1, "Base values");
    is (ref $xls,		"ARRAY",	"Return type");
    is ($xls->[0]{type},	"xlsx",		"Spreadsheet type");
    is ($xls->[0]{sheets},	2,		"Sheet count");
    is (ref $xls->[0]{sheet},	"HASH",		"Sheet list");
    is (scalar keys %{$xls->[0]{sheet}},
				2,		"Sheet list count");
    cmp_ok ($xls->[0]{version}, ">=",	0.07,	"Parser version");

    ok (1, "Defined fields");
    foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($xls->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
	is ($xls->[1]{$cell},		$cell,	"Formatted   cell $cell");
	}

    ok (1, "Undefined fields");
    foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($xls->[1]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($xls->[1]{$cell},		undef,	"Formatted   cell $cell");
	}
    my @row = Spreadsheet::Read::rows ($xls->[1]);
    is (scalar @row,		4,	"Row'ed rows");
    is (scalar @{$row[3]},	4,	"Row'ed columns");
    is ($row[0][3],		"D1",	"Row'ed value D1");
    is ($row[3][2],		"C4",	"Row'ed value C4");
    }

# Tests for empty thingies
ok ($xls = ReadData ("files/values.xlsx"), "True/False values");
ok (my $ss = $xls->[1], "first sheet");
is ($ss->{cell}[1][1],	"A1",  "unformatted plain text");
is ($ss->{cell}[2][1],	" ",   "unformatted space");
is ($ss->{cell}[3][1],	undef, "unformatted empty");
is ($ss->{cell}[4][1],	"0",   "unformatted numeric 0");
is ($ss->{cell}[5][1],	"1",   "unformatted numeric 1");
is ($ss->{cell}[6][1],	"",    "unformatted a single '");
is ($ss->{A1},		"A1",  "formatted plain text");
is ($ss->{B1},		" ",   "formatted space");
is ($ss->{C1},		undef, "formatted empty");
is ($ss->{D1},		"0",   "formatted numeric 0");
is ($ss->{E1},		"1",   "formatted numeric 1");
is ($ss->{F1},		"",    "formatted a single '");

{   # RT#74976] Error Received when reading empty sheets
    foreach my $strip (0 .. 3) {
	my $ref = ReadData ("files/blank.xlsx", strip => $strip);
	ok ($ref, "File with no content - strip $strip");
	}
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
