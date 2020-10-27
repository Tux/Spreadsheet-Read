use strict;

require 5.008001;

use ExtUtils::MakeMaker;
use File::Copy;

if ($ENV{PERLBREW_HOME} and $ENV{PERLBREW_HOME} eq "/home/njh/.perlbrew") {
     warn "Your smokers have been blocked because of consistent failures that\n";
     warn " are all caused by the smoking setup and not by module errors.  If\n";
     warn " you have fixed that all, please inform the authors, so this block\n";
     warn " can be lifted again.\n";
     exit 0;
     }

sub link_or_copy {
    my ($source, $dest) = @_;

    link ($source, $dest) or copy ($source, $dest);
    } # link_or_copy

my @exe;
unless (exists $ENV{AUTOMATED_TESTING} and $ENV{AUTOMATED_TESTING} == 1) {
    -f "examples/xlsgrep" or link_or_copy "examples/xlscat", "examples/xlsgrep";
    for ( [ "xlscat",	"Convert Spreadsheet to plain text or CSV"	],
	  [ "xlsgrep",	"Grep pattern from Spreadsheet"			],
	  [ "ss2tk",	"Show a Spreadsheet in Perl/Tk"			],
	  [ "ssdiff",	"Show diff between two spreadsheets"		],
	  [ "xls2csv",	"Wrapper around xlscat for easy XLS  => CSV"	],
	  [ "xlsx2csv",	"Wrapper around xlscat for easy XLSX => CSV"	],
	  ) {
	prompt ("Do you want to install $_->[0]\t$_->[1] ?", "y") =~ m/[Yy]/ and
	    push @exe, "examples/$_->[0]";
	}
    }
my %wm = (
    NAME         => "Spreadsheet::Read",
    DISTNAME     => "Spreadsheet-Read",
    ABSTRACT     => "Read the data from a spreadsheet",
    AUTHOR       => "H.Merijn Brand <perl5\@tux.freedom.nl>",
    VERSION_FROM => "Read.pm",
    EXE_FILES    => [ @exe ],
    PREREQ_FATAL => 0,
    PREREQ_PM    => {
	# Core modules
	"Exporter"			=> 0,
	"Encode"			=> 0,
	"Carp"				=> 0,
	"Data::Dumper"			=> 0,
	"File::Temp"			=> 0.22,
	"IO::Scalar"			=> 0,	# Optional
	"List::Util"			=> 0,

	# For testing
	"Test::More"			=> 0.88,
	"Test::NoWarnings"		=> 0,

	# for ss2tk
#	"Tk"				=> 804.032,
#	"Tk::NoteBook"			=> 4.009,
#	"Tk::TableMatrix::Spreadsheet"	=> 1.2,
	},
    macro        => { TARFLAGS   => "--format=ustar -c -v -f", },
    );
$ExtUtils::MakeMaker::VERSION > 6.30 and $wm{LICENSE} = "perl";

if ($ENV{EXTENDED_TESTING}) { # for CpanCover and masochists
    # Backend parsers, all optional
    # Versions also need to be declared inside Read.pm !
    $wm{PREREQ_PM}{"Text::CSV_PP"}		= "1.97";
    $wm{PREREQ_PM}{"Text::CSV_XS"}		= "1.36";
    $wm{PREREQ_PM}{"Spreadsheet::ReadSXC"}	= "0.26";
    $wm{PREREQ_PM}{"Spreadsheet::ParseODS"}	= "0.26";
    $wm{PREREQ_PM}{"Spreadsheet::ParseExcel"}	= "0.65";
    $wm{PREREQ_PM}{"Spreadsheet::ParseXLSX"}	= "0.27";
    $wm{PREREQ_PM}{"Spreadsheet::XLSX"}		= "0.15";
    $wm{PREREQ_PM}{"Spreadsheet::Perl"}		= "0";
    }

my $rv = WriteMakefile (%wm);

# perlcriticrc uses Config::Tiny, which does not support nesting
if (-f ".perlcriticrc" && -s "$ENV{HOME}/.perlcriticrc") {
    open my $fh, ">", ".perlcriticrc";
    print $fh do {
	local (@ARGV, $/) = ("$ENV{HOME}/.perlcriticrc"); <> };
    print $fh join "\n" => "",
	"[-Bangs::ProhibitDebuggingModules]",
	"[-BuiltinFunctions::ProhibitBooleanGrep]",
	"[-BuiltinFunctions::ProhibitStringyEval]",	# Guarded
	"[-Compatibility::PodMinimumVersion]", # L<File::Temp|https://...] requires 5.12
	"[-Documentation::ProhibitDuplicateSeeAlso]",
	"[-Documentation::RequireLinkedURLs]", # L<File::Temp|https://...] requires 5.12
	"[-Modules::ProhibitMultiplePackages]",
	"[-Subroutines::ProhibitExplicitReturnUndef]",
	"[-ValuesAndExpressions::ProhibitCommaSeparatedStatements]", # False positive
	"[-ValuesAndExpressions::ProhibitVersionStrings]",	# use 5.8.1
	"";
    close $fh;
    }

package MY;

sub postamble {
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
    my @pc = -f ".perlcriticrc" ? ("\tperlcritic -1 Read.pm") : ();
    $] >= 5.010 && -d "xt" && ($ENV{AUTOMATED_TESTING} || 0) != 1 and push @pc,
	'',
	'test::',
	'	-@env TEST_FILES="xt/*.t" make -e test_dynamic';
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
	'spellcheck:',
	'	pod-spell-check --aspell --ispell',
	'',
	'checkmeta:	spellcheck',
	'	perl sandbox/genMETA.pl -c',
	'',
	'fixmeta:	distmeta',
	'	perl sandbox/genMETA.pl',
	'	ls -l */META.yml',
	'',
	'tgzdist:	checkmeta fixmeta $(DISTVNAME).tar.gz distcheck',
	'	-@mv -f $(DISTVNAME).tar.gz $(DISTVNAME).tgz',
	'	-@cpants_lint.pl $(DISTVNAME).tgz',
	'	-@rm -f Debian_CPANTS.txt',
	'',
	'doc docs:		doc/Spreadsheet-Read.md doc/Spreadsheet-Read.html doc/Spreadsheet-Read.man',
	'	-@rm -f pod2html.tmp',
	'doc/Spreadsheet-Read.md:		Read.pm',
	'	pod2markdown  < $? > $@',
	'doc/Spreadsheet-Read.html:		Read.pm',
	'	pod2html      < $? 2>&1 | grep -v "^Cannot find" > $@',
	'doc/Spreadsheet-Read.3:		Read.pm',
	'	pod2man       < $? > $@',
	'doc/Spreadsheet-Read.man:	doc/Spreadsheet-Read.3',
	'	nroff -mandoc < $? > $@',
	@pc,
	'';
    } # postamble

1;
