#!/usr/bin/perl

use strict;
use warnings;

use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
my $parser = Spreadsheet::Read::parses ("gnumeric") or
    plan skip_all => "No Gnumeric parser found";

print STDERR "# Parser: $parser-", $parser->VERSION, "\n";

sub test_source {
    # 7 tests per file version.
    my ($source, $source_name, @options) = @_;

    ok (ref ($source) || Spreadsheet::Read::_txt_is_xml
           ($source, "http://www.gnumeric.org/v10.dtd"),
       "$source_name contains Gnumeric XML");
    my $book = ReadData ($source, @options);
    ok ($book,				"have gnumeric book");
    ok (@$book == 4,			"it has length 4");
    ok ($book->[0]{sheets} == 3,	"it has 3 sheets");
    ok (my $b1 = $book->[1],		"first sheet");
    is ($b1->{C30}, "monthly maintenance fee",
					"cell C30 matches via col/row name");
    is ($b1->{cell}[3][30], "monthly maintenance fee",
					"cell C30 matches via array indices");
}

sub test_oo_source {
    # 5 tests per file version.
    my ($source, $source_name, @options) = @_;

    my $book = Spreadsheet::Read->new ($source, @options);
    ok ($book,				"OO $source_name: have gnumeric book");
    is (scalar $book->sheets, 3,	"OO book has 3 sheets");
    ok (my $b1 = $book->sheet (1),	"OO first sheet");
    is ($b1->cell ("C30"), "monthly maintenance fee",
					"OO cell C30 matches via col/row name");
    is ($b1->cell (3, 30), "monthly maintenance fee",
					"OO cell C30 matches via array indices");
    } # test_source

### Main code.

## Source tests.

for my $file (qw(files/gnumeric.xml files/gnumeric.gnumeric)) {
    test_source ($file, "file $file");
    test_oo_source ($file, "file $file");
    open my $in, "<", $file or die "oops: $!";
    test_source ($in, "stream $file", parser => "gnumeric");
    open $in, "<", $file or die "oops: $!";
    test_oo_source ($in, "stream $file", parser => "gnumeric");
    open $in, "<", $file or die "oops again: $!";
    my $data = do { local $/; <$in> };
    test_source ($data, "scalar $file");
    test_oo_source ($data, "scalar $file");
    }

## Other basic tests.

# Check that the rc and cells flags are obeyed.
my $source = 'files/gnumeric.gnumeric';
my $book = ReadData ($source, rc => 0, cells => 1);
ok($book, "have book from $source");
is($book->[1]{B3}, 'Date', "B3 is filled in");
ok(! exists($book->[1]{cell}), "{cell} is not there");
$book = ReadData ($source, rc => 1, cells => 0);
ok($book, "have book from $source");
ok(! exists($book->[1]{B3}), "B3 is not there");
is($book->[1]{cell}[2][3], 'Date', "{cell}[2][3] is filled in");

done_testing (6 + 2 * 3 * (7 + 5));
