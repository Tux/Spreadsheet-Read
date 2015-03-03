#!/pro/bin/perl

use strict;
use warnings;

use Getopt::Long qw(:config bundling nopermute);
my $check = 0;
my $opt_v = 0;
GetOptions (
    "c|check"		=> \$check,
    "v|verbose:1"	=> \$opt_v,
    ) or die "usage: $0 [--check]\n";

use lib "sandbox";
use genMETA;
my $meta = genMETA->new (
    from    => "Read.pm",
    verbose => $opt_v,
    );

$meta->from_data (<DATA>);

if ($check) {
    $meta->check_encoding ();
    $meta->check_required ();
    my @ef = grep { !m/xls(cat|grep)|ssdiff/ } glob "examples/*";
    $meta->check_minimum ([ "t", @ef, "Read.pm", "Makefile.PL" ]);
    $meta->{h}{requires}{perl} = "5.008004";
    $meta->check_minimum ([ map { "examples/$_" } qw( xlscat xlsgrep ssdiff )]);
    $meta->done_testing ();
    }
elsif ($opt_v) {
    $meta->print_yaml ();
    }
else {
    $meta->fix_meta ();
    }

__END__
--- #YAML:1.0
name:                   Spreadsheet-Read
version:                VERSION
abstract:               Meta-Wrapper for reading spreadsheet data
license:                perl
author:                 
  - H.Merijn Brand <h.m.brand@xs4all.nl>
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
  Data::Peek:           0
  File::Temp:           0.22
configure_requires:
  ExtUtils::MakeMaker:  0
test_requires:
  Test::Harness:        0
  Test::More:           0.88
  Test::NoWarnings:     0
recommends:
  IO::Scalar:           0
  File::Temp:           0.2304
  Data::Peek:           0.43
test_recommends:
  Test::More:           1.001014
resources:
  license:              http://dev.perl.org/licenses/
  repository:           https://github.com/Tux/Spreadsheet-Read
meta-spec:
  version:              1.4
  url:                  http://module-build.sourceforge.net/META-spec-v1.4.html
optional_features:
  opt_csv:
    description:        Provides parsing of CSV streams
    requires:
      Text::CSV_XS:                        0.71
    recommends:
      Text::CSV:                           1.32
      Text::CSV_PP:                        1.31
      Text::CSV_XS:                        1.16
  opt_excel:
    description:        Provides parsing of Microsoft Excel files
    requires:
      Spreadsheet::ParseExcel:             0.34
      Spreadsheet::ParseExcel::FmtDefault: 0
    recommends:
      Spreadsheet::ParseExcel:             0.65
  opt_excelx:
    description:        Provides parsing of Microsoft Excel 2007 files
    requires:
      Spreadsheet::ParseXLSX:              0.13
      Spreadsheet::ParseExcel::FmtDefault: 0
    recommends:
      Spreadsheet::ParseXLSX:              0.16
  opt_oo:
    description:        Provides parsing of OpenOffice spreadsheets
    requires:
      Spreadsheet::ReadSXC:                0.20
  opt_tools:
    description:        Spreadsheet tools
    recommends:
      Tk:                                  804.033
      Tk::NoteBook:                        0
      Tk::TableMatrix::Spreadsheet:        0
