#!/pro/bin/perl

use 5.018002;
use warnings;

#se Excel::ValueReader::XLSX;
use Spreadsheet::Read;
#BEGIN { $ENV{SPREADSHEET_READ_XLSX} = "Excel::ValueReader::XLSX"; }

my $file = "files/Active2.xlsx";
my $book = ReadData ($file, debug => 1000, parser => "Excel::ValueReader::XLSX");
