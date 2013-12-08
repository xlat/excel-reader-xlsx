package Excel::Reader::XLSX::Package::XMLreader;

###############################################################################
#
# XMLreader - A class for reading Excel XLSX XML files.
#
# Used in conjunction with Excel::Reader::XLSX
#
# Copyright 2012, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.008002;
use strict;
use warnings;
use Exporter;
use Carp;
use XML::LibXML::Reader qw(:types);

our @ISA     = qw(Exporter);
our $VERSION = '0.00';


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;

    my $self = { _reader => undef };

    bless $self, $class;

    return $self;
}


##############################################################################
#
# _read_file()
#
# Create an XML::LibXML::Reader instance from a file.
#
sub _read_file {

    my $self     = shift;
    my $filename = shift;

    my $xml_reader = XML::LibXML::Reader->new(
        location  => $filename,
        no_blanks => 1
    );

    $self->{_reader} = $xml_reader;
    $self->{_source} = { location => $filename };

    return $xml_reader;
}


##############################################################################
#
# _read_string()
#
# Create an XML::LibXML::Reader instance from a string. Used mainly for
# testing.
#
sub _read_string {

    my $self   = shift;
    my $string = shift;

    my $xml_reader = XML::LibXML::Reader->new(
        string    => $string,
        no_blanks => 1
    );

    $self->{_reader} = $xml_reader;
    $self->{_source} = { string    => $string };

    return $xml_reader;
}


##############################################################################
#
# _read_filehandle()
#
# Create an XML::LibXML::Reader instance from a filehandle. Used mainly for
# testing.
#
sub _read_filehandle {

    my $self       = shift;
    my $filehandle = shift;

    my $xml_reader = XML::LibXML::Reader->new(
        IO        => $filehandle,
        no_blanks => 1
    );

    $self->{_reader} = $xml_reader;
    #not useful, it could not be cloned, may require to wrap $filehandle in something
    #that allow to rewind stream (by storing data in a tmp file?)
    $self->{_source} = { IO => $filehandle };

    return $xml_reader;
}


##############################################################################
#
# _read_all_nodes()
#
# Read all the nodes of an Excel XML file using an XML::LibXML::Reader
# instance. Sub-classes will provide the _read_node() method.
#
sub _read_all_nodes {

    my $self = shift;

    while ( $self->{_reader}->read() ) {
        $self->_read_node( $self->{_reader} );
    }
}


##############################################################################
#
# _parse_file()
#
# Shortcut for the most common use case: _read_file() + _read_all_nodes().
#
sub _parse_file {

    my $self     = shift;
    my $filename = shift;

    my $xml_reader = $self->_read_file( $filename );
    $self->_read_all_nodes();

    return $xml_reader;
}

##############################################################################
#
# clone()
#
# Clone the reader and return another one that will restart document from the begining.
#
sub clone {
    my $self = shift;
    my $clone = __PACKAGE__->new;

    die "clone not implemented on filehandle!" 
        if exists $self->{_source}->{IO};

   my $xml_reader = XML::LibXML::Reader->new(
        %{$self->{_source}},
        no_blanks => 1
    );

    $clone->{_reader} = $xml_reader;
    $clone->{_source} = $self->{_source};

    return $clone;
}

#not tested, that's just a draft...
# It will be possible to "fork" by reading cloned until it reach the same file position that current object.
# It may not works if moveTo* methods was called...
sub fork{
    my $self = shift;
    my $fork = $self->clone;
    my $bytes = $self->{_reader}->byteConsumed;
    my $freader = $fork->{_reader};
    while($bytes && $freader->read){
        last if $freader->byteConsumed == $bytes;
    }
    return $fork;
}


#Rewind the reader to the begining
##############################################################################
#
# rewind()
#
# Rewind the reader so it will restart document from the begining.
# To be true, it build a new one.
#
sub rewind{
    my $self = shift;
    my $clone = $self->clone;
    $self->{_reader} = $clone->{_reader};
}
1;


__END__

=pod

=head1 NAME

XMLreader - A class for reading Excel XLSX XML files.

=head1 SYNOPSIS

See the documentation for L<Excel::Reader::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Reader::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Reader::XLSX>.

=cut
