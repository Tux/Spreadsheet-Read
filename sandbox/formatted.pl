#!/pro/bin/perl

use 5.18.2;
use warnings;

use Spreadsheet::Read;

my $file     = "files/example.xlsx";
my $workbook = Spreadsheet::Read->new ($file);

my $info     = $workbook->[0];
say "Parsed $file with $info->{parser}-$info->{version}";

my $sheet    = $workbook->sheet (1);

say join "\t" => "Formatted:",   $sheet->row     (1);
say join "\t" => "Unformatted:", $sheet->cellrow (1);
