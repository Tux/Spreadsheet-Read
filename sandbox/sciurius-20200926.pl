#!/pro/bin/perl

use 5.014002;
use warnings;

our $VERSION = "0.01 - 20200927";
our $CMD = $0 =~ s{.*/}{}r;

sub usage {
    my $err = shift and select STDERR;
    say "usage: $CMD ...";
    exit $err;
    } # usage

use CSV;
use Spreadsheet::Read;
use Getopt::Long qw(:config bundling);
GetOptions (
    "help|?"		=> sub { usage (0); },
    "V|version"		=> sub { say "$CMD [$VERSION]"; exit 0; },

    "v|verbose:1"	=> \(my $opt_v = 0),
    ) or usage (1);

my $ss = Spreadsheet::Read->new ("sciurius-20200926.ods", attr => 1);

DDisplay $ss->sheet (1)->cell ("D8");
DDisplay $ss->sheet (1)->cell (4, 8);
DDumper $ss->[1]{attr}[4][8];
