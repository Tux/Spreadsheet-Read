#!/pro/bin/perl

use 5.22.0;
use warnings;

use Data::Peek;
use Spreadsheet::XLSX::Reader::LibXML;

my $p = Spreadsheet::XLSX::Reader::LibXML->new;
my $b = $p->parse ("files/merged.xlsx");

DDumper $b;

my $s = 0;
foreach my $w ($b->worksheets) {

    say "SHEET ", ++$s;
    DDumper $w;
    my ($row_min, $row_max) = $w->row_range;
    my ($col_min, $col_max) = $w->col_range;

    foreach my $row ($row_min .. $row_max) {
	foreach my $col ($col_min .. $col_max) {
	    my $cell = $w->get_cell ($row, $col) or next;

	    print "Row, Col    = ($row, $col)\n";
	    print "Value       = ", $cell->value(),       "\n";
	    print "Unformatted = ", $cell->unformatted(), "\n";
	    print "\n";
	    }
	}
    }

#DDumper $wb;

