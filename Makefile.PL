use strict;

require 5.006;

use ExtUtils::MakeMaker;

my @exe;
for ( [ "xlscat",	"Convert Spreadsheet to plain text or CSV"	],
      [ "ss2tk",	"Show a Spreadsheet in Perl/Tk"			],
      ) {
    prompt ("Do you want to install '$_->[0]' ($_->[1]) ? ", "n") =~ m/[Yy]/ and
	push @exe, "examples/$_->[0]";
    }
my %wm = (
    NAME         => "Spreadsheet::Read",
    ABSTRACT     => "Read the data from a spreadsheet",
    AUTHOR       => "H.Merijn Brand <h.merijn\@procura.nl>",
    VERSION_FROM => "Read.pm",
    EXE_FILES    => [ @exe ],
    PREREQ_FATAL => 0,
    PREREQ_PM    => {
# Backend parsers
	"Text::CSV_XS"			=> 0.00,
	"Text::CSV_PP"			=> 1.05,
	"Spreadsheet::ReadSXC"		=> 0.20,
	"Spreadsheet::ParseExcel"	=> 0.26,
#	"Spreadsheet::Perl"		=> 0.00,	# Not yet

	# Core modules
	"File::Temp"			=> 0.14,

# for ss2tk
#	"Tk"				=> 804.027,
#	"Tk::NoteBook"			=> 4.009,
#	"Tk::TableMatrix::Spreadsheet"	=> 1.2,
	},
    );
$ExtUtils::MakeMaker::VERSION > 6.30 and $wm{LICENSE} = "perl";

my $rv = WriteMakefile (%wm);

package MY;

sub postamble
{
    join "\n" =>
	'cover test_cover:',
	'	cover -delete',
	'	HARNESS_PERL_SWITCHES=-MDevel::Cover $(MAKE) test',
	'	cover',
	'';
    } # postamble

1;
