use 5.012000;
use warnings;
use Spreadsheet::Read;
use Data::Peek;
use JSON;

our $VERSION = "0.01";

BEGIN { *SR:: = \%Spreadsheet::Read::; }

sub sr  { Spreadsheet::Read->new (@_);	}
sub dsr { DDumper     (sr (@_));	}
sub jsr { decode_json (sr (@_));	}

1;

=head1 NAME

SR - Alias for Spreadheet::Read

=head1 SYNOPSIS

  perl -MSR -wE'say sr ("file.xls")->sheet (1)->cell ("B3")'
  perl -MSR -wE'dsr ("file.xlsx")'

  perl -MSR -wE'DDumper (csv (in => \q{foo,"bar, foo",quux}))'
  perl -MSR -wE'dcsv (in => \q{foo,"bar, foo",quux})'

  perl -MSR -wE'print jsr ("file.ods")'

=head1 DESCRIPTION

Wrapper for Text::CSV_XS with csv importing also Data::Peek and JSON
See L<Text::CSV_XS>, L<Data::Peek>, L<JSON>

=head2 sr

Alias for C<< Spreadsheet::Read->new >>

=head2 dsr

Uses L<Data::Peek>'s DDumper to output the result of C<sr>

=head2 jsr

Uses L<JSON>'s decode_json to convert the result of C<sr> to JSON.

=head1 AUTHOR

H.Merijn Brand <hmbrand@cpan.org>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2020-2024 H.Merijn Brand

This library is free software; you can redistribute it and/or modify it
under the same terms as Perl itself.

=cut
