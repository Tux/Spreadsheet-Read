#!/pro/bin/perl

# ss-dup-tk.pl: Find dups in spreadsheet
#	  (m)'09 [23-01-2009] Copyright H.M.Brand 2005-2014

use strict;
use warnings;

sub usage
{
    my $err = shift and select STDERR;
    print
	"usage: $0 [-t] [-S <sheets>] [-R <rows>] [-C columns] [-F <fields>]\n",
	"\t-t          Only check on true values\n",
	"\t-S sheets   Check sheet(s).  Defaul = 1,   1,3-5,all\n",
	"\t-R rows     Check row(s).    Defaul = all, 6,19-66\n",
	"\t-C columns  Check column(s). Defaul = all, 2,5-9\n",
	"\t-F fields   Check field(s).  Defaul = all, A1,A2,B15,C23\n";
    exit $err;
    } # usage

use Spreadsheet::Read;

use Getopt::Long qw(:config bundling nopermute noignorecase);
my $opt_v = 0;
my $opt_t = 0;		# Only check on true values
my @opt_S;		# Sheets to print
my @opt_R;		# Rows to print
my @opt_C;		# Columns to print
my @opt_F;
GetOptions (
    "help|?"		=> sub { usage (0); },

    "S|sheets=s"	=> \@opt_S,
    "R|rows=s"		=> \@opt_R,
    "C|columns=s"	=> \@opt_C,
    "F|fields=s"	=> \@opt_F,

    "t|true"		=> \$opt_t,

    "v|verbose:1"	=> \$opt_v,
    ) or usage (1);

@opt_S or @opt_S = (1);


use Tk;
use Tk::ROText;

my $file = shift || (sort { -M $b <=> -M $a } glob "*.xls")[0];
my ($mw, $is, $ss, $dt) = (MainWindow->new, "1.0");

sub ReadFile
{
    $file or return;

    $dt->delete ("1.0", "end");
    unless ($ss = ReadData ($file)) {
	$dt->insert ("end", "Cannot read $file as spreadsheet\n");
	return;
	}

    my @ss = map { qq{"$ss->[$_]{label}"} } 1 .. $ss->[0]{sheets};

    my @finfo = (
	"File: $file", ( map {
	    "Sheet $_: '$ss->[$_]{label}'\t($ss->[$_]{maxcol} x $ss->[$_]{maxrow})"
	    } 1 .. $ss->[0]{sheets} ),
	"==============================================================");
    $dt->insert ("end", join "\n", @finfo, "");
    $is = (@finfo + 1).".0";
    return $ss;
    } # ReadFile

my $tf = $mw->Frame ()->pack (qw( -side top -anchor nw -expand 1 -fill both ));
$tf->Entry (
    -textvariable	=> \$file,
    -width		=> 40,
    -vcmd		=> \&ReadFile,
    )->pack (qw(-side left -expand 1 -fill both));

my %ftyp;
for ([ xls  => [ "Excel Files",		[qw( .xls  .XLS  )] ] ],
     [ xlsx => [ "Excel Files",		[qw( .xlsx .XLSX )] ] ],
     [ sxc  => [ "OpenOffice Files",	[qw( .sxc  .SXC  )] ] ],
     [ ods  => [ "OpenOffice Files",	[qw( .ods  .ODS  )] ] ],
     [ csv  => [ "CSV Files",		[qw( .csv  .CSV  )] ] ],
     ) {
    my ($ft, $r) = @$_;
    Spreadsheet::Read::parses ($ft) or next;
    push @{$ftyp{$r->[0]}}, @{$r->[1]};
    push @{$ftyp{"All spreadsheet types"}}, @{$r->[1]};
    }
$tf->Button (
    -text	=> "Select file",
    -command	=> sub {
	$ss   = undef;
	$file = $mw->getOpenFile (
	    -filetypes => [ ( map { [ $_, $ftyp{$_} ] } sort keys %ftyp ),
		[ "All files", "*" ],
		],
	    );
	ReadFile ();
	},
    )->pack (qw(-side left -expand 1 -fill both));
$tf->Button (
    -text	=> "Detect",
    -command	=> \&Detect,
    )->pack (qw(-side left -expand 1 -fill both));
$tf->Button (
    -text	=> "Show",
    -command	=> \&Show,
    )->pack (qw(-side left -expand 1 -fill both));
$tf->Button (
    -text	=> "Exit",
    -command	=> \&exit,
    )->pack (qw(-side left -expand 1 -fill both));

my $mf = $mw->Frame  ()->pack (qw( -side top -anchor nw -expand 1 -fill both ));
my $sw = $mf->Scrolled ("ROText",
    -scrollbars             => "osoe",
	-height             => 40,
	-width              => 85,
	-foreground         => "Black",
	-background         => "White",
	-highlightthickness => 0,
	-setgrid            => 1)->pack (qw(-expand 1 -fill both));
$dt = $sw->Subwidget ("scrolled");
#$sw->Subwidget ("xscrollbar")->packForget;
$dt->configure (
    -wrap	=> "none",
    -font	=> "mono 12",
    );

my $bf = $mw->Frame  ()->pack (qw( -side top -anchor nw -expand 1 -fill both ));
$bf->Checkbutton (
    -variable	=> \$opt_t,
    -text	=> "True values only",
    )->pack (qw(-side left -expand 1 -fill both));
{   my $opt_S = @opt_S ? join ",", @opt_S : 1;
    $bf->Label (
	-text		=> "Sheet(s)",
	)->pack (qw(-side left -expand 1 -fill both));
    $bf->Entry (
	-textvariable	=> \$opt_S,
	-width		=> 10,
	-validate	=> "focusout",
	-vcmd		=> sub {
	    @opt_S = grep m/\S/, split m/\s*,\s*/ => $opt_S;
	    1;
	    },
	)->pack (qw(-side left -expand 1 -fill both));
    }
{   my $opt_R = join ",", @opt_R;
    $bf->Label (
	-text		=> "Rows(s)",
	)->pack (qw(-side left -expand 1 -fill both));
    $bf->Entry (
	-textvariable	=> \$opt_R,
	-width		=> 10,
	-validate	=> "focusout",
	-vcmd		=> sub {
	    @opt_R = grep m/\S/, split m/\s*,\s*/ => $opt_R;
	    1;
	    },
	)->pack (qw(-side left -expand 1 -fill both));
    }
{   my $opt_C = join ",", @opt_C;
    $bf->Label (
	-text		=> "Columns(s)",
	)->pack (qw(-side left -expand 1 -fill both));
    $bf->Entry (
	-textvariable	=> \$opt_C,
	-width		=> 10,
	-validate	=> "focusout",
	-vcmd		=> sub {
	    @opt_C = grep m/\S/, split m/\s*,\s*/ => $opt_C;
	    1;
	    },
	)->pack (qw(-side left -expand 1 -fill both));
    }

sub ranges (@)
{
    my @g;
    foreach my $arg (@_) {
	for (split m/,/, $arg) {
	    if (m/^(\w+)\.\.(\w+)$/) {
		my ($s, $e) = ($1, $2);
		$s =~ m/^[1-9]\d*$/ or ($s, $e) = (qq("$s"), qq("$e"));
		eval "push \@g, $s .. $e";
		}
	    else {
		push @g, $_;
		}
	    }
	}
    $opt_v and print STDERR "( @g )\n";
    @g;
    } # ranges

sub Detect
{
    $ss or ReadFile ();

    $dt->delete ($is, "end");
    $dt->insert ("end", join "\n", "",
	"Shts: @opt_S",
	"Rows: @opt_R",
	"Cols: @opt_C",
	"--------------------------------------------------------------",
	"");
    my %done;
    my @S = $opt_S[0] eq "all" ? (1 .. $ss->[0]{sheets}) : ranges (@opt_S);
    my @R = ranges (@opt_R);
    my @C = ranges (@opt_C);
    my %f = map { uc $_ => 1 } ("@opt_F" =~ m/(\b[A-Z]\d+\b)/ig);

    foreach my $s (@S) {
	my $xls = $ss->[$s] or die "Cannot read sheet $s\n";

	my @r = @R ? @R : (1 .. $xls->{maxrow});
	my @c = @C ? @C : (1 .. $xls->{maxcol});

	foreach my $r (@r) {
	    foreach my $c (@c) {
		defined $xls->{cell}[$c][$r] or next;
		my $v    = uc $xls->{cell}[$c][$r];
		my $cell = cr2cell ($c, $r);
		@S > 1 and $cell = $xls->{label} . "[$cell]";

		$opt_t && !$v		and next;
		@opt_F && !exists $f{$cell} and next;

		if (exists $done{$v}) {
		    $dt->insert ("end", sprintf "Cell %-5s is dup of %-5s '%s'\n", $cell, $done{$v}, $v);
		    next;
		    }
		$done{$v} = $cell;
		}
	    }
	}
    } # Detect

sub Show
{
    $ss or ReadFile ();

    $dt->delete ($is, "end");
    $dt->insert ("end", join "\n", "",
	"Shts: @opt_S",
	"Rows: @opt_R",
	"Cols: @opt_C");
    my @S = $opt_S[0] eq "all" ? (1 .. $ss->[0]{sheets}) : ranges (@opt_S);
    my @R = ranges (@opt_R);
    my @C = ranges (@opt_C);
    my %f = map { uc $_ => 1 } ("@opt_F" =~ m/(\b[A-Z]\d+\b)/ig);

    foreach my $s (@S) {
	my $xls = $ss->[$s] or die "Cannot read sheet $s\n";

	$dt->insert ("end",
	    "\n--------------------------------------------------------------".
	    "\nSheet $s: '$xls->{label}'\t($xls->{maxcol} x $xls->{maxrow})\n");

	my @r = @R ? @R : (1 .. $xls->{maxrow});
	my @c = @C ? @C : (1 .. $xls->{maxcol});

	$dt->insert ("end", "     |");
	for (@c) {
	    (my $ch = cr2cell ($_, 1)) =~ s/1$//;
	    $dt->insert ("end", sprintf "%11s |", $ch);
	    }
	$dt->insert ("end", "\n-----+");
	$dt->insert ("end", "------------+") for @c;
	foreach my $r (@r) {
	    $dt->insert ("end", sprintf "\n%4d |", $r);
	    foreach my $c (@c) {
		my $cell = cr2cell ($c, $r);
		my $v = defined $xls->{cell}[$c][$r]
		    ? $xls->{$cell}
		    : "--";
		length ($v) < 12 and substr $v, 0, 0, " " x (12 - length $v);
		$dt->insert ("end", substr ($v, 0, 12). "|");
		}
	    }
	}
    } # Show

MainLoop;
