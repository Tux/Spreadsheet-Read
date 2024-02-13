#!/usr/bin/perl

use strict;
use warnings;

use     Test::More;
require Test::NoWarnings;

use Spreadsheet::Read;
my $parser;
if ($parser = Spreadsheet::Read::parses ("xls")) {
    plan tests => 191;
    Test::NoWarnings->import;
    }
else {
    plan skip_all => "No M\$-Excel parser found";
    }

sub ReadDataStream {
    my $file = shift;
    open my $fh, "<", $file or return undef;
    ReadData ($fh, parser => "xls", @_);
    } # ReadDataStream

{   my $ref;
    $ref = ReadDataStream ("no_such_file.xls");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadDataStream ("files/empty.xls");
    ok (!defined $ref, "Empty file");
    }

my $content;
{   local $/;
    open my $xls, "<", "files/test.xls" or die "files/test.xls: $!";
    binmode $xls;
    $content = <$xls>;
    close   $xls;
    }

my $xls;
foreach my $base ( [ "files/test.xls",	"Read/Parse xls file"	],
#		   [ $content,		"Parse xls data"	],
		   ) {
    my ($txt, $msg) = @$base;
    ok ($xls = ReadDataStream ($txt),	$msg);

    ok (1, "Base values");
    is (ref $xls,		"ARRAY",	"Return type");
    is ($xls->[0]{type},	"xls",		"Spreadsheet type");
    is ($xls->[0]{sheets},	2,		"Sheet count");
    is (ref $xls->[0]{sheet},	"HASH",		"Sheet list");
    is (scalar keys %{$xls->[0]{sheet}},
				2,		"Sheet list count");
    cmp_ok ($xls->[0]{version}, ">=",	0.26,	"Parser version");

    ok (1, "Defined fields");
    foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($xls->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
	is ($xls->[1]{$cell},		$cell,	"Formatted   cell $cell");
	}

    ok (1, "Undefined fields");
    foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($xls->[1]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($xls->[1]{$cell},		undef,	"Formatted   cell $cell");
	}
    my @row = Spreadsheet::Read::rows ($xls->[1]);
    is (scalar @row,		4,	"Row'ed rows");
    is (scalar @{$row[3]},	4,	"Row'ed columns");
    is ($row[0][3],		"D1",	"Row'ed value D1");
    is ($row[3][2],		"C4",	"Row'ed value C4");
    }

# This files is generated under Mac OS/X Tiger
ok (1, "XLS File fom Mac OS X");
ok ($xls = ReadDataStream ("files/macosx.xls", clip => 0),
						"Read/Parse Mac OS X xls file");

ok (1, "Base values");
is ($xls->[0]{sheets},		3,		"Sheet count");
is ($xls->[0]{sheet}{Sheet3},	3,		"Sheet labels");
is ($xls->[1]{maxrow},		25,		"MaxRow");
is ($xls->[1]{maxcol},		3,		"MaxCol");
is ($xls->[2]{label},		"Sheet2",	"Sheet label");
is ($xls->[2]{maxrow},		0,		"Empty sheet maxrow");
is ($xls->[2]{maxcol},		0,		"Empty sheet maxcol");

ok (1, "Content");
is ($#{$xls->[1]{cell}[3]}, $xls->[1]{maxrow}, "cell structure");
ok (defined $xls->[1]{cell}[$xls->[1]{maxcol}][$xls->[1]{maxrow}], "last cell");

foreach my $x (1 .. 17) {
    my $cell = cr2cell (1, $x);
    is ($xls->[1]{$cell},		$x,	"Cell $cell");
    is ($xls->[1]{cell}[1][$x],		$x,	"Cell 1, $x");
    }
foreach my $x (1 .. 25) {
    my $cell = cr2cell (3, $x);
    is ($xls->[1]{$cell},		$x,	"Cell $cell");
    is ($xls->[1]{cell}[3][$x],		$x,	"Cell 3, $x");
    }
foreach my $cell (qw( A18 B1 B6 B20 C26 D14 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($xls->[1]{cell}[$c][$r],	undef,	"Cell $cell");
    is ($xls->[1]{$cell},		undef,	"Cell $c, $r");
    }

eval {
    eval "use ".$parser."::FmtDefault";
    my ($pm) = map { $INC{$_} } grep m{FmtDefault.pm$}i => keys %INC;
    if (open my $ph, "<", $pm) {
	my $l;
	$l = <$ph> for 1 .. 68;
	close $ph;
	if ($l =~ m/'C\*'/) {
	    print STDERR "\n",
			 "# If the next tests give warnings like\n",
			 "# Character in 'C' format wrapped in pack at\n",
			 "#    $pm line 68\n",
			 "# Change C* to U* in line 68\n",
			 "# patch -p0 <SPE68.diff\n";
	    my @patch = <DATA>;
	    s/\bPM\b/$pm/ for @patch;
	    open  $ph, ">", "SPE68.diff" or die "SPE68.diff: $!\n";
	    print $ph @patch;
	    close $ph;
	    }
	}
    };

# Tests for empty thingies
ok ($xls = ReadDataStream ("files/values.xls"), "True/False values");
ok (my $ss = $xls->[1], "first sheet");
is ($ss->{cell}[1][1],	"A1",  "unformatted plain text");
is ($ss->{cell}[2][1],	" ",   "unformatted space");
is ($ss->{cell}[3][1],	undef, "unformatted empty");
is ($ss->{cell}[4][1],	"0",   "unformatted numeric 0");
is ($ss->{cell}[5][1],	"1",   "unformatted numeric 1");
is ($ss->{cell}[6][1],	"",    "unformatted a single '");
is ($ss->{A1},		"A1",  "formatted plain text");
is ($ss->{B1},		" ",   "formatted space");
is ($ss->{C1},		undef, "formatted empty");
is ($ss->{D1},		"0",   "formatted numeric 0");
is ($ss->{E1},		"1",   "formatted numeric 1");
is ($ss->{F1},		"",    "formatted a single '");

# Test extended attributes (active & hidden)
ok ($xls = Spreadsheet::Read->new ("files/Active2.xls", attr => 1), "Active sheet");
is ($xls->sheets,		3,	"Book has 3 sheets");
SKIP: {
    my $v = $xls->[0]{version};
    $v < 0.61 and skip "$xls->[0]{parser}-$v does not support the active flag", 3;
    is ($xls->sheet (1)->{active},	0, "Sheet 1 is not active");
    is ($xls->[2]{active},		1, "Sheet 2 is active");
    is ($xls->sheet (3)->active,	0, "Sheet 3 is not active");
    is ($xls->sheet (1)->{hidden},	0, "Sheet 1 is not hidden");
    is ($xls->[2]{hidden},		0, "Sheet 2 is not hidden");
    is ($xls->sheet (3)->hidden,	0, "Sheet 3 is not hidden");
    }

is ($xls->sheet (1)->attr ("A1")->{type}, "text", "Attr through method A1");
is ($xls->sheet (1)->attr (2, 2)->{type}, "text", "Attr through method B2");

ok ($xls = Spreadsheet::Read->new ("files/attr.xls", attr => 1), "Attributes OO");
is ($xls->[1]{attr}[3][3]{fgcolor},		"#008000", "C3 Forground color direct");
is ($xls->sheet (1)->attr (3, 3)->{fgcolor},	"#008000", "C3 Forground color OO rc   hash");
is ($xls->sheet (1)->attr ("C3")->{fgcolor},	"#008000", "C3 Forground color OO cell hash");
is ($xls->sheet (1)->attr (3, 3)->fgcolor,	"#008000", "C3 Forground color OO rc   method");
is ($xls->sheet (1)->attr ("C3")->fgcolor,	"#008000", "C3 Forground color OO cell method");

is ($xls->[1]{attr}[3][3]{bogus_attribute},	undef, "C3 bogus attribute direct");
is ($xls->sheet (1)->attr ("C3")->{bogus_attr},	undef, "C3 bogus attribute OO hash");
is ($xls->sheet (1)->attr ("C3")->bogus_attr,	undef, "C3 bogus attribute OO method");

__END__
--- PM    2005-09-15 14:16:36.163623616 +0200
+++ PM    2005-09-15 14:11:56.289171000 +0200
@@ -65,7 +65,7 @@ sub new($;%) {
 sub TextFmt($$;$) {
     my($oThis, $sTxt, $sCode) =@_;
     return $sTxt if((! defined($sCode)) || ($sCode eq '_native_'));
-    return pack('C*', unpack('n*', $sTxt));
+    return pack('U*', unpack('n*', $sTxt));
 }
 #------------------------------------------------------------------------------
 # FmtStringDef (for Spreadsheet::ParseExcel::FmtDefault)
