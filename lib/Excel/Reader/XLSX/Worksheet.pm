package Excel::Reader::XLSX::Worksheet;

###############################################################################
#
# Worksheet - A class for reading the Excel XLSX sheet.xml file.
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
use Carp;
use Excel::Reader::XLSX::Package::XMLreader;
use Excel::Reader::XLSX::Row;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';
our $USE_CACHE = 1;

###############################################################################
#
# Public and private API methods.
#
###############################################################################


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $self  = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_shared_strings}      = shift;
    $self->{_name}                = shift;
    $self->{_index}               = shift;
    $self->{_previous_row_number} = -1;
    $self->{row_cache} = [] if $USE_CACHE;

    bless $self, $class;

    return $self;
}

###############################################################################
#
# get_link( $range )
#
# Return an hash reference if the requested $range has an hyperlink.
# The hash contain the following keys: location, display
#
sub get_link{
        my ($self, $range) = @_;
        $self->_init_link unless exists $self->{_links};
        return $self->{_links}->{ $range };
}

###############################################################################
#
# follow_link( $link )
#
# Return the cell of the hyperlink target in scalar context. 
# Return $worksheet, $row, $cell in list context 
# It is cross sheet but not (YET) cross workbook.
#
sub follow_link{
        my ($self, $link, $wantsheet, $wantbook) = @_;
        $wantsheet //= 1;
        if($link->{location} and $link->{location} =~ /^(?|'([^']+)'!(.*)|([^!]+)!(.*))$/){
            my ($sheet, $range) = ($1, $2);
            my $worksheet = $self->{_book}->worksheet( $sheet );
            return scalar $worksheet->get_range($range) unless wantarray;
            return ( $worksheet->get_range($range, $wantsheet, $wantbook) );
        }
}

###############################################################################
#
# get_range( $range )
#
# In scalar context, return the Cell object that match $range or undef if it doesn't exists.
# In list context, return the Row and Cell object that match $range or undef if it doesn't exists.
#
sub get_range{
    my ($self, $range, $wantsheet, $wantbook) = @_;
    my @sub_ranges = $self->{_book}->parse_range( $range ) or return;
    my ($book_name, $sheet_name, $row_number, $cols, $subrange) = @{$sub_ranges[0]};
    $sheet_name //= $self->name;
    my $sheet = $self->name eq $sheet_name ? $self : $self->{_book}->worksheet($sheet_name);
    my $row = $sheet->get_row( $row_number );
    my $cell = $row->get_cell( $cols );
    return $cell unless wantarray;
    return ( $row, $cell ) unless $wantsheet or $wantbook;
    return ( $sheet, $row, $cell ) if $wantsheet and not $wantbook;
    my $reader = Excel::Reader::XLSX->new();
    my $book = $reader->read_file( $book_name ) 
        or die $reader->error(), "\n";
    return ( $book, ($wantsheet ? $sheet : undef), $row, $cell ) if $wantbook;
}



###############################################################################
#
# set_row()
#
# Set the next available row in the worksheet (only available when $USE_CACHE is true).
#
sub set_row{
    my $self = shift;
    my $row  = shift // ($self->{_previous_row_number} + 1);
    die "set_row cannot be called without \$USE_CACHE!" unless $USE_CACHE;
    die "set_row called with $row but only " . scalar(@{$self->{row_cache}}) . " row cached!" if $row > @{$self->{row_cache}};
    $self->{_previous_row_number} = $row;
    $self->{_row} = $self->{row_cache}[$row];
}


###############################################################################
#
# next_row()
#
# Read the next available row in the worksheet.
#
sub next_row {

    my $self = shift;
    my $row  = undef;

    if($USE_CACHE and ($self->{_previous_row_number}>=0 or @{$self->{row_cache}}) ){
        while($self->{_previous_row_number} < @{$self->{row_cache}} -1){
            my $row_obj = $self->set_row;
            return $row_obj if ref $row_obj;
        }
    }

    # Read the next "row" element in the file.
    return unless $self->{_reader}->nextElement( 'row' );

    # Read the row attributes.
    my $row_reader = $self->{_reader};
    my $row_number = $row_reader->getAttribute( 'r' );

    # Zero index the row number.
    if ( defined $row_number ) {
        $row_number--;
    }
    else {

        # If no 'r' attribute assume it is one more than the previous.
        $row_number = $self->{_previous_row_number} + 1;
    }

    if ( !$self->{_row_initialised} or $USE_CACHE ) {
        $self->_init_row();
    }

    $row = $self->{_row};
    $row->_init( $row_number );
    $self->{_previous_row_number} = $row_number;
    
    if($USE_CACHE){
        $self->{row_cache}[$row_number]=$row;
    }

    return $row;
}

###############################################################################
#
# get_row( $row_number )
#
# return the Row object that match $row_number or undef if it doesn't exists.
#
sub get_row{
    my ($self, $row_number) = @_;
    die "called with inconsistant row: $row_number" if $row_number < 0;
    if($USE_CACHE){
        if($row_number < @{$self->{row_cache}} - 1){
            return $self->set_row( $row_number );
        }
    }
    elsif($row_number < $self->{_previous_row_number}){
        $self->rewind;
        $self->{_previous_row_number} = -1;
    }
    my $row = $self->{_row};
    while( ($self->{_previous_row_number} < $row_number)
             and ($row = $self->next_row) ){

    }
    return $row;
}

###############################################################################
#
# name()
#
# Return the worksheet name.
#
sub name {

    my $self = shift;

    return $self->{_name};
}


###############################################################################
#
# index()
#
# Return the worksheet index.
#
sub index {

    my $self = shift;

    return $self->{_index};
}


###############################################################################
#
# Internal methods.
#
###############################################################################

#Overload of the Rewind the reader to the begining
##############################################################################
#
# rewind()
#
# Rewind the reader so it will restart document from the begining.
# To be true, it build a new one.
#
sub rewind{
    my $self = shift;
    if($USE_CACHE){
        $self->{_previous_row_number} = -1;
        for(@{$self->{row_cache}}){
            $_->{_next_cell_index} = 0 if ref;
        }
        return $self->{_reader};
    }
    else{
        return $self->SUPER::rewind;
    }
}

###############################################################################
#
# _init_row()
#
# TODO.
#
sub _init_row {

    my $self = shift;

    # Store reusable Cell object to avoid repeated calls to Cell::new().
    $self->{_cell} = Excel::Reader::XLSX::Cell->new( $self, $self->{_shared_strings} );

    # Store reusable Row object to avoid repeated calls to Row::new().
    $self->{_row}  = Excel::Reader::XLSX::Row->new(
        $self,
        $self->{_shared_strings},
        $self->{_cell},
    );

    $self->{_row_initialised} = 1;
}

###############################################################################
#
# _init_link( )
#
# Read all hyperlinks and store them as an hash reference under $self->{_links}
#
sub _init_link{
        my $self = shift;
        # Set up the file to read.
        my $reader = $self->clone->{_reader};
        my %links;
        if($reader->nextElement('hyperlinks')){
            my $link_node = $reader->copyCurrentNode( 1 );
            my @hyperlink_nodes = $link_node->getChildrenByTagName( 'hyperlink' );
            foreach(@hyperlink_nodes){
                my $ref_range = $_->getAttribute('ref');
                my %target = ( 
                                location => $_->getAttribute('location'), 
                                display  => $_->getAttribute('display')
                            );
               foreach my $ref ( $self->{_book}->parse_range( $ref_range ) ){
                    $links{ $ref->[-1] } = \%target;
                }
            }
        }
        $self->{_links} = \%links;
}
###############################################################################
#
# _init( $workbook, $sheetprops )
#
# Initialize current Worksheet object with it's $workbook and $sheetprops, so it doesn't need external
# manipulation to build XML reader on demand. (eg: _init_link )
#
sub _init{
        my $self = shift;
        $self->{_book} = shift;
        $self->{_props} = shift;    # from workbook _worksheet_properties
        my $filename =  $self->{_book}->{_package_dir}
                        . $self->{_book}->{_workbook_root}
                        . $self->{_props}->{_filename};

    # Set up the file to read. We don't read data until it is required.
    $self->_read_file( $filename );
}

1;


__END__

=pod

=head1 NAME

Worksheet - A class for reading the Excel XLSX sheet.xml file.

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
