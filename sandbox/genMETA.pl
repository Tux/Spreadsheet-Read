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

my $version;
open my $pm, "<", "Read.pm" or die "Cannot read Read.pm";
while (<$pm>) {
    m/^our\s+.VERSION\s*=\s*"?([-0-9._]+)"?\s*;\s*$/ or next;
    $version = $1;
    last;
    }
close $pm;

my @yml;
while (<DATA>) {
    s/VERSION/$version/o;
    push @yml, $_;
    }

if ($check) {
    use YAML::Syck;
    use Test::YAML::Meta::Version;
    my $h;
    my $yml = join "", @yml;
    eval { $h = Load ($yml) };
    $@ and die "$@\n";
    $opt_v and print Dump $h;
    my $t = Test::YAML::Meta::Version->new (yaml => $h);
    $t->parse () and die join "\n", $t->errors, "";

    use Parse::CPAN::Meta;
    eval { Parse::CPAN::Meta::Load ($yml) };
    $@ and die "$@\n";

    my $req_vsn = $h->{requires}{perl};
    print "Checking if $req_vsn is still OK as minimal version\n";
    use Test::MinimumVersion;
    all_minimum_version_ok ($req_vsn, { paths =>
	["t", "examples", "Read.pm", "Makefile.PL" ]});
    }
elsif ($opt_v) {
    print @yml;
    }
else {
    my @my = glob <*/META.yml>;
    @my == 1 && open my $my, ">", $my[0] or die "Cannot update META.yml\n";
    print $my @yml;
    close $my;
    chmod 0644, $my[0];
    }

__END__
--- #YAML:1.1
name:                   Read
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
recommends:
  perl:                 5.010001
  File::Temp:           0.22
  IO::Scalar:           0
configure_requires:
  ExtUtils::MakeMaker:  0
build_requires:
  perl:                 5.006
  Test::Harness:        0
  Test::More:           0.88
  Test::NoWarnings:     0
resources:
  license:              http://dev.perl.org/licenses/
  repository:           http://repo.or.cz/w/Spreadsheet-Read.git
meta-spec:
  version:              1.4
  url:                  http://module-build.sourceforge.net/META-spec-v1.4.html
optional_features:
- opt_csv:
    description:        Provides parsing of CSV streams
    requires:
      Text::CSV_XS:                        0.69
    recommends:
      Text::CSV:                           1.15
      Text::CSV_PP:                        1.23
      Text::CSV_XS:                        0.69
- opt_excel:
    description:        Provides parsing of Microsoft Excel files
    requires:
      Spreadsheet::ParseExcel:             0.26
      Spreadsheet::ParseExcel::FmtDefault: 0
    recommends:
      Spreadsheet::ParseExcel:             0.55
- opt_excelx:
    description:        Provides parsing of Microsoft Excel 2007 files
    requires:
      Spreadsheet::XLSX:                   0.12
      Spreadsheet::XLSX::Fmt2007:          0
- opt_oo:
    description:        Provides parsing of OpenOffice spreadsheets
    requires:
      Spreadsheet::ReadSXC:                0.2
- opt_tools:
    description:        Spreadsheet tools
    recommends:
      Tk:                                  0
      Tk::NoteBook:                        0
      Tk::TableMatrix::Spreadsheet:        0
