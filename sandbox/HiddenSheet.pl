#!/pro/bin/perl

use 5.016003;
use warnings;

use Data::Peek;
use Spreadsheet::Read;

my $file = shift;
my $book = ReadData ($file, dtfmt => "DD.MM.YYYY", attr => 1);
say "after readdata";

my $xls_info = $book->[0];
my $sheets   = $xls_info->{sheets};

say "excel file sheets $sheets";
$sheets or die "no sheets $sheets";

my $worksheet = 0;
if ($sheets > 1) {
    foreach my $sheet (1 .. $sheets) {
	my $item = $book->[$sheet];
	#delete $item->{attr};delete $item->{cell}; DDumper $item;
	$item->{hidden} and say "HIDDEN $sheet";
	say "SheetHidden $item->{hidden} $sheet";
	if ($item->{active}) {
	    $worksheet = $sheet;    # found an active
	    last;
	    }
	}
    unless ($worksheet) {           # hat nur ein sheet zeilen?
	say "multiple worksheets not one active: $sheets sheets: @{[keys %{$xls_info->{sheet}}]}";
	foreach my $sheet (1 .. $sheets) {
	    my $item   = $book->[$sheet];
	    my $maxrow = 0;
	    $maxrow = $item->{maxrow} if defined $item->{maxrow};
	    say "worksheets: $sheet maxrow $maxrow";
	    if ($maxrow) {
		$worksheet and die "multiple sheets have rows";
		$worksheet = $sheet;
		}
	    }
	}
    }
else {
    $worksheet = $sheets;
    }
