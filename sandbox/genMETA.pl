#!/pro/bin/perl

use strict;
use warnings;

my $version;
open my $pm, "<", "Read.pm" or die "Cannot read Read.pm";
while (<$pm>) {
    m/^our\s+.VERSION\s*=\s*"?([-0-9._]+)"?\s*;\s*$/ or next;
    $version = $1;
    last;
    }
close $pm;

my @my = glob <*/META.yml>;
@my == 1 && open my $my, ">", $my[0] or die "Cannot update META.yml|n";
while (<DATA>) {
    s/VERSION/$version/o;
    print $my $_;
    }
close $my;

__END__
--- #YAML:1.0
name:                   Read
version:                VERSION
abstract:               Meta-Wrapper for reading spreadsheet data
license:                perl
author:                 
  - H.Merijn Brand <h.merijn@xs4all.nl>
generated_by:           Author
distribution_type:      module
provides:
  Spreadsheet::Read:
    file:               Read.pm
    version:            VERSION
requires:                       
  perl:                 5.006
  Exporter:             0
  Carp:                 0
  Data::Dumper:         0
recommends:
  File::Temp:           0.14
  IO::Scalar:           0
build_requires:
  perl:                 5.006
  Test::Harness:        0
  Test::More:           0

optional_features:
  - CSV:
    description:        Provides parsing of CSV streams
    requires:
      Text::CSV_XS:     0.23
    recommends:
      Text::CSV:        1
      Text::CSV_PP:     1.05
      Text::CSV_XS:     0.52
  - Excel:
    description:        Provides parsing of Microsoft Excel files
    requires:
      Spreadsheet::ParseExcel: 0.26
    recommends:
      Spreadsheet::ParseExcel: 0.33
  - OpenOffice:
    description:        Provides parsing of OpenOffice spreadsheets
    requires:
      Spreadsheet::ReadSXC:    0.2
  - Tools:
    description:        Spreadsheet tools
    recommends:
      Tk:                           0
      Tk::NoteBook:                 0
      Tk::TableMatrix::Spreadsheet: 0
resources:
  license:      http://dev.perl.org/licenses/
meta-spec:
  url:          http://module-build.sourceforge.net/META-spec-v1.3.html
  version:      1.3
