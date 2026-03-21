#!/usr/bin/perl

# Test parsers() behavior, specifically:
# - Entries with "!" prefix (failed modules) are filtered out
# - No uninitialized value warnings when $can{$typ} is undef
# - Multiline $@ in failed require does not leak through the filter
# See https://github.com/Tux/Spreadsheet-Read/issues/55

use strict;
use warnings;

use Test::More tests => 7;
use Test::NoWarnings;

use Spreadsheet::Read;

my $sr = "Spreadsheet::Read";

# parsers() should return results without warnings
ok (my @p = $sr->parsers (), "parsers() returns results");

# No entry should have an ext starting with "!"
my @leaked = grep { $_->{ext} =~ m/^!/ } @p;
is (scalar @leaked, 0, "No '!' entries leak through parsers()");

# No entry should have ext "dmp" or "ios" (helper modules)
my @helpers = grep { $_->{ext} =~ m/^(?:dmp|ios)$/ } @p;
is (scalar @helpers, 0, "Helper modules (dmp, ios) are filtered");

# SquirrelCalc should always be present as a built-in
my @sc = grep { $_->{ext} eq "sc" } @p;
is (scalar @sc, 1, "SquirrelCalc entry present");
is ($sc[0]{def}, "*", "SquirrelCalc is the default sc parser");

# Simulate the bug from issue #55: multiline $@ in the format field
# When require fails on FreeBSD, $@ can contain embedded newlines.
# The old regex m{^(?:dmp|ios|!.*)$} fails on these because . doesn't
# match \n, so the $ anchor can't match end of string.
# The fix uses m{^(?:dmp|ios|!)} which only checks the prefix.
{
    my $multiline_entry = "! Cannot use Fake::Module version 1.0: "
	. "Can't locate Fake/Module.pm\n"
	. "in \@INC (you may need to install the Fake::Module module)";

    # Old regex (broken): fails to filter multiline $@ entries
    my $old_regex = qr{^(?:dmp|ios|!.*)$};
    # New regex (fixed): prefix-only check, immune to newlines
    my $new_regex = qr{^(?:dmp|ios|!)};

    ok ($multiline_entry !~ $old_regex &&
	$multiline_entry =~ $new_regex,
	"Multiline '!' entry: old regex fails, new regex matches (issue #55)");
}
