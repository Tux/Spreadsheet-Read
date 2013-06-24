#!/pro/bin/perl

use v5.14;
use warnings;

use Data::Peek;
use Spreadsheet::ParseExcel;

my $p = Spreadsheet::ParseExcel->new (); #Password => "kitchen7");
my $s = $p->parse ("TAYLORS2_PMPINTA-APP.xls") or
    die $p->error ();

foreach my $w ($s->worksheets ()) {
    print "Sheet: $w->{Name}\n";
    local $" = "..";
    print " rows: @{[$w->row_range()]}\n";
    print " cols: @{[$w->col_range()]}\n";
    }
