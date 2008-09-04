use strict;

require 5.006;

use ExtUtils::MakeMaker;

my @exe;
unless (exists $ENV{AUTOMATED_TESTING} and $ENV{AUTOMATED_TESTING} == 1) {
    for ( [ "xlscat",	"Convert Spreadsheet to plain text or CSV"	],
	  [ "ss2tk",	"Show a Spreadsheet in Perl/Tk"			],
	  [ "xls2csv",	"Wrapper around xlscat for easy XLS => CSV"	],
	  ) {
	prompt ("Do you want to install '$_->[0]' ($_->[1]) ? ", "n") =~ m/[Yy]/ and
	    push @exe, "examples/$_->[0]";
	}
    }
my %wm = (
    NAME         => "Spreadsheet::Read",
    ABSTRACT     => "Read the data from a spreadsheet",
    AUTHOR       => "H.Merijn Brand <h.merijn\@procura.nl>",
    VERSION_FROM => "Read.pm",
    EXE_FILES    => [ @exe ],
    PREREQ_FATAL => 0,
    PREREQ_PM    => {
	# Core modules
	"Exporter"			=> 0.00,
	"Data::Dumper"			=> 0.00,
	"File::Temp"			=> 0.14,
	"IO::Scalar"			=> 0.00,	# Optional

	# Backend parsers, all optional
	"Text::CSV"			=> 1.00,
	"Text::CSV_XS"			=> 0.23,	# 0.32 or up preferred
	"Text::CSV_PP"			=> 1.05,	# 1.08 or up preferred
	"Spreadsheet::ReadSXC"		=> 0.20,
	"Spreadsheet::ParseExcel"	=> 0.26,
#	"Spreadsheet::Perl"		=> 0.00,	# Not yet

	# For testing
	"Test::More"			=> 0.00,

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
    my $valgrind = join " ", qw(
	PERL_DESTRUCT_LEVEL=2 PERL_DL_NONLAZY=1
	    valgrind
		--suppressions=sandbox/perl.supp
		--leak-check=yes
		--leak-resolution=high
		--show-reachable=yes
		--num-callers=50
		--log-fd=3
		$(FULLPERLRUN) "-MExtUtils::Command::MM" "-e"
		    "test_harness($(TEST_VERBOSE), '$(INST_LIB)', '$(INST_ARCHLIB)')"
		    $(TEST_FILES) 3>valgrind.log
		    );
    join "\n" =>
	'cover test_cover:',
	'	ccache -C',
	'	cover -test',
	'',
	'leakcheck:',
	"	$valgrind",
	'	-@tail -5 valgrind.log',
	'',
	'leaktest:',
	q{	sandbox/leaktest $(FULLPERLRUN) "test_harness($(TEST_VERBOSE), '$(INST_LIB)', '$(INST_ARCHLIB)')" $(TEST_FILES)},
	'',
	'checkmeta:',
	'	perl sandbox/genMETA.pl -c',
	'',
	'fixmeta:	distmeta',
	'	perl sandbox/genMETA.pl',
	'	ls -l */META.yml',
	'',
	'tgzdist:	checkmeta fixmeta $(DISTVNAME).tar.gz distcheck',
	'	-@mv -f $(DISTVNAME).tar.gz $(DISTVNAME).tgz',
	'	-@cpants_lint.pl $(DISTVNAME).tgz';
    } # postamble

1;
