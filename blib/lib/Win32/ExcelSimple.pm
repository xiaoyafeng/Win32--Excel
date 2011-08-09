package Win32::ExcelSimple;

use 5.010;
use Moose;
use Array::Transpose::Ragged qw/transpose_ragged/;
use Spreadsheet::ConvertAA;
use Exporter 'import'; # gives you Exporter's import() method directly
our @EXPORT_OK = qw(transpose_array aa2cell cell2aa);  # symbols to export on request
use Moose::Util::TypeConstraints;
use Spreadsheet::Read;

subtype 'xls_file',
      as 'Str',
      where { $_ =~ /xls$/ },
      message { "please choose xls file!!\n" };

=head1 NAME

Win32::ExcelSimple - The great new Win32::ExcelSimple!

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';
has 'file_name' => (is => 'ro', isa => 'xls_file', required => 1);
has 'write_obj' => (is => 'ro', isa => 'Spreadsheet::Read');
has 'read_obj'  => (is => 'ro', isa => 'Win32::OLE');

=head1 SYNOPSIS

Quick summary of what the module does.

Perhaps a little code snippet.

    use Win32::ExcelSimple;

    my $foo = Win32::ExcelSimple->new();
    ...

=head1 EXPORT

A list of functions that can be exported.  You can delete this section
if you don't export anything, such as for a purely object-oriented module.

=head1 SUBROUTINES/METHODS

=head2 transpose_array


=cut
no Moose;
sub transpose_array {

my @trans = transpose_ragged($_[0]);
return \@trans;
}

=head2 cell2aa

=cut

sub cell2aa {
	return ToAA(shift);
}


=head2 aa2cell

=cut

sub aa2cell {
	return FromAA(shift);

}





=head1 AUTHOR

Andy Xiao, C<< <andy.xiao at gmail.com> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-win32-excelsimple at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=Win32-ExcelSimple>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.




=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc Win32::ExcelSimple


You can also look for information at:

=over 4

=item * RT: CPAN's request tracker (report bugs here)

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=Win32-ExcelSimple>

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/Win32-ExcelSimple>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/Win32-ExcelSimple>

=item * Search CPAN

L<http://search.cpan.org/dist/Win32-ExcelSimple/>

=back


=head1 ACKNOWLEDGEMENTS


=head1 LICENSE AND COPYRIGHT

Copyright 2011 Andy Xiao.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.


=cut

1; # End of Win32::ExcelSimple
