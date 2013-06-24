#!/pro/bin/perl

use v5.14.1;
use warnings;

use Data::Peek;
use Spreadsheet::Read;

my $ss2read = "TAYLORS2_PMPINTA-APP.xls";

-f $ss2read or die "$ss2read: no such file";

#my $ref1 = ReadData ("$ss2read");
my $ref1 = ReadData ($ss2read, debug => 9);#, Password => "kitchen7");

print "A10: $ref1->[1]{A10}\n";    # content of field A10 of sheet 1

print "Row 10: (", (join ", " => map { $_ ? DDisplay ($_) : "--undef--" }
    Spreadsheet::Read::row ($ref1->[1], 10)), ")\n";

my @rows = Spreadsheet::Read::rows ($ref1->[1]);

for my $i (0 .. $#rows) {
    for my $j (0 .. $#{$rows[$i]}) {
	defined (my $v = $rows[$i][$j]) or next;
	printf "[ %3d, %3d ] %s\n", $i, $j, DDisplay ($v);
	}
    }

print "Finished.\n";
