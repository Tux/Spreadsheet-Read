#!/pro/bin/perl

use 5.014002;
use warnings;

use Spreadsheet::Read;

my $book = ReadData ("issue34.xlsx") or die "couldn't open Excel sheet\n";
my $meta = $book->[0];
my $data = $book->[1];
say "Using Spreadsheet::Read-", Spreadsheet::Read->VERSION;
say "With $meta->{type} parser $meta->{parser}-$meta->{version}";
my $tmp_str = $data->{cell}[1][1];
my @tmp_arr = split m/^/ => $tmp_str;
say "'$tmp_str'";
@tmp_arr > 1
    ? say "PASSED"
    : die "FAILED\n please use sim-4 or maxl05\n"; 
