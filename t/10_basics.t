#!/usr/bin/perl

use strict;
use warnings;

use Test::More tests => 118;
use Test::NoWarnings;

use Spreadsheet::Read qw(:DEFAULT parses rows );

is (Spreadsheet::Read::Version (), $Spreadsheet::Read::VERSION, "Version check");

ok (my @ext = parses (undef), "No sheet type");
is_deeply ([ grep m/^sc/ => @ext ], [ "sc" ], "Supports Squirrelcalc");

is (parses ("zzz0"), 0,				"Unknown     sheet type");
is (parses ("zzz1"), 0,				"Too old     sheet type");
is (parses ("zzz2"), "Z20::Just::For::Testing",	"Testing     sheet type");
is (parses ("zzz3"), 0,				"Unsupported sheet type");

is (parses ("xls"), parses ("excel"),      "Excel alias type");
is (parses ("ods"), parses ("oo"),         "OpenOffice alias type 1");
is (parses ("ods"), parses ("OpenOffice"), "OpenOffice alias type 2");
is (parses ("sc"),  parses ("scalc"),      "SquirrelCalc alias type");

foreach my $x ([ "A1",              1,      1 ],
               [ "Z26",            26,     26 ],
               [ "AB12",           28,     12 ],
               [ "A",               0,      0 ],
               [ "19",              0,      0 ],
               [ "a9",              1,      9 ],
               [ "aAa9",          703,      9 ],
               [ "",                0,      0 ],
               [ undef,             0,      0 ],
               [ "x444444",        24, 444444 ],
               [ "xxxxxx4", 296559144,      4 ],
               ) {
    my $cell = $x->[0];
    my ($c, $r) = cell2cr ($x->[0]);
    defined $cell or $cell = "";
    is ($c, $x->[1], "Col for $cell");
    is ($r, $x->[2], "Row for $cell");
    }

foreach my $x ([         1,      1, "A",      "A1"      ],
               [        26,     26, "Z",      "Z26"     ],
               [        28,     12, "AB",     "AB12"    ],
               [         0,      0, "",       ""        ],
               [        -2,      0, "",       ""        ],
               [         0,    -12, "",       ""        ],
               [         1,    -12, "A",      ""        ],
               [     undef,      1, "",       ""        ],
               [         2,  undef, "B",      ""        ],
               [         1,      9, "A",      "A9"      ],
               [       703,      9, "AAA",    "AAA9"    ],
               [        24, 444444, "X",      "X444444" ],
               [ 296559144,      4, "XXXXXX", "XXXXXX4" ],
               ) {
    my $cell = cr2cell ($x->[0], $x->[1]);
    my ($c, $r) = map { defined $_ ? $_ : "--undef--" } $x->[0], $x->[1];
    is ($cell, $x->[3], "Cell for ($c, $r)");

    is (Spreadsheet::Read::col2label ($x->[0]), $x->[2], "Col $c eq $x->[2]");
    }

# Some illegal rows () calls
for (undef, "", " ", 0, 1, [], {}) {
    my @rows = rows ($_);
    my $arg = defined $_ ? $_ : "-- undef --";
    is (scalar @rows, 0, "Illegal rows ($arg)");
    }
for (undef, "", " ", 0, 1, [], {}) {
    my @rows = rows ({ cell => $_});
    my $arg = defined $_ ? $_ : "-- undef --";
    is (scalar @rows, 0, "Illegal rows ({ cell => $arg})");
    }
for (undef, "", " ", 0, 1, [], {}) {
    my @rows = rows ({ maxrow => 1, cell => $_});
    my $arg = defined $_ ? $_ : "-- undef --";
    is (scalar @rows, 0, "Illegal rows ({ maxrow => 1, cell => $arg })");
    }
for (undef, "", " ", 0, 1, [], {}) {
    my @rows = rows ({ maxcol => 1, cell => $_});
    my $arg = defined $_ ? $_ : "-- undef --";
    is (scalar @rows, 0, "Illegal rows ({ maxcol => 1, cell => $arg })");
    }

# Some illegal ReadData () calls
for (undef, "", " ", 0, 1, [], {}) {
    my $ref = ReadData ($_);
    my $arg = defined $_ ? $_ : "-- undef --";
    is ($ref, undef, "Illegal ReadData ($arg)");
    }
for (undef, "", " ", 0, 1, [], {}) {
    my $ref = ReadData ([ $_ ]);
    my $arg = defined $_ ? $_ : "-- undef --";
    is ($ref, undef, "Illegal ReadData ([ $arg ])");
    }
SKIP: {
    -c "/dev/null" or skip "/dev/null cannot be used for tests", 7;
    for (undef, "", " ", 0, 1, [], {}) {
	my $ref = ReadData ("/dev/null", separator => $_);
	my $arg = defined $_ ? $_ : "-- undef --";
	is ($ref, undef, "Illegal ReadData ({ $arg })");
	}
    }
for (undef, "", " ", 0, 1, [], {}) {
    my $ref;
    eval { $ref = ReadData ("Read.pm", sep => $_); };
    my $arg = defined $_ ? $_ : "-- undef --";
    is ($ref, undef, "Illegal ReadData ({ $arg })");
    }

my $sr = "Spreadsheet::Read";
ok (my @p = $sr->parsers (), "Parser info");
is_deeply ((grep { $_->{ext} eq "sc" } @p)[0],
    { ext => "sc", min => "0.01", mod => $sr, vsn => $sr->VERSION, def => "*" },
    "SquirrelCalc is internal");
