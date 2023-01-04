#!/pro/bin/perl

use 5.012000;
use warnings;

use Data::Peek;
use Spreadsheet::Read;
use encoding "cp936", STDOUT => "cp936";

my $ss = ReadData ("xiao.xls")->[2];

DDumper { A3 => DDisplay ($ss->{A3}), 21 => DDisplay ($ss->{cell}[2][1]) };
say $ss->{A3};
say $ss->{cell}[2][1];

