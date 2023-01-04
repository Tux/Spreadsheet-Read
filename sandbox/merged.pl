#!/pro/bin/perl

use 5.020000;
use warnings;
use Spreadsheet::Read;

my $ss = ReadData ("merged.xlsx", attr => 1, debug => 99)->[1];

foreach my $row (1 .. $ss->{maxrow}) {
    foreach my $col (1 .. $ss->{maxcol}) {
	my $cell = cr2cell ($col, $row);
	printf "%s %-6s %d  ", $cell, $ss->{$cell},
	    $ss->{attr}[$col][$row]{merged};
	}
    print "\n";
    }

say for keys %$ss;
foreach my $area (@{$ss->{merged}}) {
    print cr2cell (@{$area}[0,1]), " -> ", cr2cell (@{$area}[2,3]), "\n";
    }
