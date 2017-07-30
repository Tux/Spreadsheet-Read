#!/pro/bin/perl

use 5.18.2;
use warnings;
use Data::Peek;
use Spreadsheet::Read;

my $ss = ReadData ("rt114903-fail.xlsx", attr => 1);

DDumper $ss->[0];

__END__
#!/usr/bin/perl

#--------------------------------------------------------------------------------------------------
# early.pl by snzieg 20160316
#--------------------------------------------------------------------------------------------------

$|=1;
use strict;
use warnings;

use Data::Peek;
use Spreadsheet::Read;

open  DOUT, ">", "dumper.LOG";      # Falls debug, dann File anlegen
print DOUT DDumper ("dumper.LOG File created");

my $INFILE = "File Early Ad Haircare_error";

say "\nreading $INFILE.xlsx";
my $ReadXLSX = ReadData ("$INFILE.xlsx", attr => 1) or die "Fehler: $!\n";

print DOUT DDumper ($ReadXLSX);
close DOUT;
