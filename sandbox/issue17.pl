#!/pro/bin/perl

use 5.18.2;
use warnings;

use Data::Peek;
use Spreadsheet::Read;

my $dir = -d "files" ? "files" : "../files";
open my $fh, "<", "$dir/test.ods";
DDumper ReadData ($fh, parser => "ods");
close $fh;
