#!/usr/bin/perl

use strict;
use warnings;

use     Test::More;
require Test::NoWarnings;

use     Spreadsheet::Read;
my $parser = Spreadsheet::Read::parses ("gnumeric") or
    plan skip_all => "No Gnumeric parser found";

print STDERR "# Parser: $parser-", $parser->VERSION, "\n";

# Subroutine.

sub test_source {
    # 6 tests per file version.
    my ($source, $source_name, @options) = @_;

    ok(ref($source) ||
       Spreadsheet::Read::_txt_is_xml
           ($source, 'http://www.gnumeric.org/v10.dtd'),
       "$source_name contains Gnumeric XML");
    my $book = ReadData($source, @options);
    ok($book, "have gnumeric book");
    ok(@$book == 4, "it has length 4");
    ok($book->[0]{sheets} == 3, "it has 3 sheets");
    ok($book->[1]{C30} eq 'monthly maintenance fee',
       'cell C30 matches via col/row name');
    ok($book->[1]{cell}[3][30] eq 'monthly maintenance fee',
       'cell C30 matches via array indices');
}

### Main code.

for my $file (qw(files/gnumeric.xml files/gnumeric.gnumeric)) {
    test_source ($file, "file $file");
    open my $in, '<', $file or die "oops: $!";
    test_source ($in, "stream $file", parser => 'gnumeric');
    open $in, '<', $file or die "oops again: $!";
    my $data = do { local $/;  <$in> };
    test_source ($data, "scalar $file");
}

done_testing(2 * 3 * 6);
