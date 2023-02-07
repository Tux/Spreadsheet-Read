#!/usr/bin/perl

use 5.012000;
use diagnostics; ## verbose errors
use warnings FATAL => qw( all           );
use Encode            qw( encode decode );
use Spreadsheet::Read qw( ReadData      );
use Data::Peek;					# Instead of Data::Dumper

my $workbook = Spreadsheet::Read->new ("accent.xlsx");
# DDumper { workbook => $workbook };
say "Parsed with $workbook->[0]{parser}-$workbook->[0]{version}";

binmode STDOUT, ":encoding(utf-8)";
binmode STDERR, ":encoding(utf-8)";

my $sheet = $workbook->sheet (1);
my $cell  = $sheet->cell ("A1");

DPeek ($cell);
say "cell A1 raw  = '$cell'";
my $no_utfed = encode ("UTF-8", $cell);
DPeek ($no_utfed);
say "cell A1 enc  = '$no_utfed'";
my $utfed    = decode ("UTF-8", $cell);
DPeek ($utfed);
say "cell A1 dec  = '$utfed'";
