# General

I am always open to improvements and suggestions. Use
[issues](https://github.com/Tux/Spreadsheet-Read/issues)

# Style

I will never accept pull request that do not strictly conform to my
style, however you might hate it. You can read the reasoning behind
my [preferences](http://tux.nl/style.html).

I really do not care about mixed spaces and tabs in (leading) whitespace

Perl::Tidy will help getting the code in shape, but as all software, it
is not perfect. You can find my preferences for these in
[.perltidy](https://github.com/Tux/Release-Checklist/blob/master/.perltidyrc) and
[.perlcritic](https://github.com/Tux/Release-Checklist/blob/master/.perlcriticrc).

# Mail

Please, please, please, do *NOT* use HTML mail.
[Plain text](https://useplaintext.email)
[without](http://www.goldmark.org/jeff/stupid-disclaimers/)
[disclaimers](https://www.economist.com/business/2011/04/07/spare-us-the-e-mail-yada-yada)
will do fine!

# Requirements

The minimum version required to use this module is stated in
[Makefile.PL](./Makefile.PL).  That does however not guarantee that it will work
for all underlying parsers, as they might require newer perl versions.

# New parsers

I am all open to support new parsers. The closer the API is to that of
Spreadsheet::ParseExcel, the easier it will be to support it.
