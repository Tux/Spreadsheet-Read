#!/pro/bin/perl

use 5.18.2;
use warnings;

# http://www.perlmonks.org/?node_id=1209625

use Spreadsheet::Read;

my $ss = Spreadsheet::Read->new ("pm1209625.xlsx", debug => 9, verbose => 9);
