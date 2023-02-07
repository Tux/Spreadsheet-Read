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
$meta->gen_cpanfile ();

if ($check) {
    $meta->check_encoding ();
    $meta->check_required ();
    my @ef = grep { !m/xls(cat|grep)|ssdiff/ } map { glob "$_/*" } qw( examples scripts );
    $meta->check_minimum ([ "t", @ef, "Read.pm", "Makefile.PL" ]);
    $meta->{h}{requires}{perl} = "5.008004";
    $meta->check_minimum ([ map { "scripts/$_" } qw( xlscat xlsgrep )]);
    $meta->{h}{requires}{perl} = "5.14";
    $meta->check_minimum ([ map { "scripts/$_" } qw( ssdiff )]);
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
  - H.Merijn Brand <perl5@tux.freedom.nl>
generated_by:           Author
distribution_type:      module
provides:
  Spreadsheet::Read:
    file:               Read.pm
    version:            VERSION
requires:                       
  perl:                 5.008001
  Exporter:             0
  Carp:                 0
  Data::Dumper:         0
  Data::Peek:           0
  Encode:               0
  File::Temp:           0.22
  List::Util:           0
configure_requires:
  ExtUtils::MakeMaker:  0
test_requires:
  Test::Harness:        0
  Test::More:           0.88
  Test::NoWarnings:     0
recommends:
  IO::Scalar:           0
  File::Temp:           0.2311
  Data::Peek:           0.52
test_recommends:
  Test::More:           1.302191
resources:
  license:              http://dev.perl.org/licenses/
  repository:           https://github.com/Tux/Spreadsheet-Read
  bugtracker:           https://github.com/Tux/Spreadsheet-Read/issues
meta-spec:
  version:              1.4
  url:                  http://module-build.sourceforge.net/META-spec-v1.4.html
optional_features:
  opt_csv:
    description:        Provides parsing of CSV streams
    requires:
      Text::CSV_XS:                        0.71
    recommends:
      Text::CSV:                           2.02
      Text::CSV_PP:                        2.02
      Text::CSV_XS:                        1.49
  opt_xls:
    description:        Provides parsing of Microsoft Excel files
    requires:
      Spreadsheet::ParseExcel:             0.34
      Spreadsheet::ParseExcel::FmtDefault: 0
    recommends:
      Spreadsheet::ParseExcel:             0.65
  opt_xlsx:
    description:        Provides parsing of Microsoft Excel 2007 files
    requires:
      Spreadsheet::ParseXLSX:              0.24
      Spreadsheet::ParseExcel::FmtDefault: 0
    recommends:
      Spreadsheet::ParseXLSX:              0.27
  opt_ods:
    description:        Provides parsing of OpenOffice spreadsheets
    requires:
      Spreadsheet::ParseODS:               0.26
    recommends:
      Spreadsheet::ParseODS:               0.37
  opt_sxc:
    description:        Provides parsing of OpenOffice spreadsheets old style
    requires:
      Spreadsheet::ReadSXC:                0.26
    recommends:
      Spreadsheet::ReadSXC:                0.37
  opt_gnumeric:
    description:        Provides parsing of Gnumeric spreadsheets
    requires:
      Spreadsheet::ReadGnumeric:           0.1
    recommends:
      Spreadsheet::ReadGnumeric:           0.1
  opt_tools:
    description:        Spreadsheet tools
    recommends:
      Tk:                                  804.036
      Tk::NoteBook:                        0
      Tk::TableMatrix::Spreadsheet:        0
