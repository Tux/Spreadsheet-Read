#!/pro/bin/perl

use 5.18.0;
use warnings;
use Spreadsheet::Read;

my $book = ReadData ("julia-good.xlsx");
say "Book was parsed using $book->[0]{parser}-$book->[0]{version}";
say "A1: ", $book->[1]{A1}; 
my @rows = Spreadsheet::Read::rows ($book->[1]);
foreach my $i (1 .. scalar @rows) {
    foreach my $j (1 .. scalar @{$rows[$i-1]}) {
        say chr (64 + $i) . " $j " . ($rows[$i - 1][$j - 1] // "");
	}
    }
