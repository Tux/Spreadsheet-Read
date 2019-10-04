#!/pro/bin/perl

use 5.18.3;
use warnings;

use Data::Peek;
use HTTP::Tiny;
use Spreadsheet::Read;

# Fetch data and return a filehandle to that data
sub fh_from_url {
    my $url = shift;
    my $ua  = HTTP::Tiny->new;
    my $res = $ua->get ($url);
    open my $fh, "<", \$res->{content};
    return $fh
    } # fh_from_url

my $fh = fh_from_url ("http://tux.nl/Files/dist-perl.csv");
my $sheet = Spreadsheet::Read->new ($fh, parser => "csv");
DDumper $sheet->[1];
