#!/pro/bin/perl

use 5.016002;
use warnings;
use autodie;

use Text::CSV_XS;

open my $fh, "<", $ARGV[0];

my $header;
read $fh, $header, 12;	# Really no idea

my @data;

# 0402 vvvv	- vvvv is length of field + control
# rrrr cccc ???? llll
my $buf;
while (read $fh, $buf, 4) {

    my ($h0, $fldlen) = unpack "vv", $buf;

    $fldlen or next;

    read $fh, $buf, $fldlen;

    # x was always 0
    my ($row, $col, $x, $val) = unpack "vv v v/a", $buf;

    $data[$row][$col] = $val;
    }

my $csv = Text::CSV_XS->new ({ binary => 1, auto_diag => 1, eol => "\n" });
$csv->print (*STDOUT, $_) for @data;
