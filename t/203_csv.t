#!/usr/bin/perl

use strict;
use warnings;

my     $tests = 125;
use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
my $parser = Spreadsheet::Read::parses ("csv") or
    plan skip_all => "No CSV parser found";

sub ReadDataStream {
    my $file = shift;
    open my $fh, "<", $file or return undef;
    ReadData ($fh, parser => "csv", @_);
    } # ReadDataStream

{   my $ref;
    $ref = ReadDataStream ("no_such_file.csv");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadDataStream ("files/empty.csv");
    ok (!defined $ref, "Empty file");
    }

my $csv;
ok ($csv = ReadDataStream ("files/test.csv"),	"Read/Parse csv file");

ok (1, "Base values");
is (ref $csv,			"ARRAY",	"Return type");
is ($csv->[0]{type},		"csv",		"Spreadsheet type");
is ($csv->[0]{sheets},		1,		"Sheet count");
is (ref $csv->[0]{sheet},	"HASH",		"Sheet list");
is (scalar keys %{$csv->[0]{sheet}},
				1,		"Sheet list count");
cmp_ok ($csv->[0]{version},	">=",	0.01,	"Parser version");

is ($csv->[1]{maxrow},		5,		"Last row");
is ($csv->[1]{maxcol},		19,		"Last column");
is ($csv->[1]{cell}[$csv->[1]{maxcol}][$csv->[1]{maxrow}],
				"LASTFIELD",	"Last field");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		$cell,	"Formatted   cell $cell");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	"",   	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		"",   	"Formatted   cell $cell");
    }

ok ($csv = ReadDataStream ("files/test_m.csv", sep => ";"),	"Read/Parse M\$ csv file");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		$cell,	"Formatted   cell $cell");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	"",   	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		"",   	"Formatted   cell $cell");
    }

ok ($csv = ReadDataStream ("files/test_x.csv", sep => "=", quote => "_"),
					    "Read/Parse strange csv file");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		$cell,	"Formatted   cell $cell");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	"",   	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		"",   	"Formatted   cell $cell");
    }

my $data = "a,b,c\n1,2,3\n";
ok ($csv = ReadData ( $data, parser => "csv"), "Parse from plain string");
is ($csv->[1]{C2}, 3, "C2 = 3");
ok ($csv = ReadData (\$data, parser => "csv"), "Parse from ref to plain string");
is ($csv->[1]{C2}, 3, "C2 = 3");
open my $rh, "<", \$data;
ok ($csv = ReadData ($rh,    parser => "csv"), "Parse from data-ref");
is ($csv->[1]{C2}, 3, "C2 = 3");

if (open my $dh, ">", $data) {
    print $dh $data;
    close $dh;
    ok ($csv = ReadData ($data, parser => "csv"), "Parse from file with bad name");
    is ($csv->[1]{C2}, 3, "C2 = 3");
    unlink $data;
    }
else {
    ok (1, "Using a badly named file not allowed: GOOD!");
    ok (1, "Using a badly named file not allowed: GOOD!");
    }

unless ($ENV{AUTOMATED_TESTING}) {
    Test::NoWarnings::had_no_warnings ();
    $tests++;
    }
done_testing ($tests);
