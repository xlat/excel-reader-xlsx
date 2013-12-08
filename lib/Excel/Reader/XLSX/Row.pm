package Excel::Reader::XLSX::Row;

###############################################################################
#
# Row - A class for reading Excel XLSX rows.
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
use XML::LibXML::Reader;
use Excel::Reader::XLSX::Cell;
use Excel::Reader::XLSX::Package::XMLreader;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';

our $FULL_DEPTH = 1;


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $self  = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_sheet}         = shift;
    $self->{_shared_strings} = shift;
    $self->{_cell}           = shift;

    bless $self, $class;

    return $self;
}


###############################################################################
#
# _init()
#
# TODO.
#
sub _init {

    my $self = shift;

    $self->{_row_number}          = shift;
    $self->{_previous_row_number} = $self->{_sheet}->{_previous_row_number};
	$self->{_reader}              = $self->{_sheet}->{_reader};
    $self->{_row_is_empty}        = $self->{_reader}->isEmptyElement();
    $self->{_values}              = undef;

    # TODO. Make the cell initialisation a lazy load.
    # Read the child cell nodes.
    my $row_node   = $self->{_reader}->copyCurrentNode( $FULL_DEPTH );
    my @cell_nodes = $row_node->getChildrenByTagName( 'c' );
				
    $self->{_cells}               = \@cell_nodes;
    $self->{_max_cell_index}  = scalar @cell_nodes;
    $self->{_next_cell_index} = 0;
}


###############################################################################
#
# next_cell()
#
# Get the next cell in the current row.
#
sub next_cell {
    my $self = shift;
		
    return if $self->{_row_is_empty};

    return if $self->{_next_cell_index} >= $self->{_max_cell_index};

    my $cell = $self->_mk_cell( $self->{_next_cell_index} );

    $self->{_next_cell_index}++;

    return $cell;
}

###############################################################################
#
# get_cell( $col_index )
#
# Get the cell at $col_index in current Row object.
#
sub get_cell{
    my ($self, $col_index) = @_;
    #TODO: improve performance by caching columns indexes.
    for my $col_node_idx (0 .. $self->{_max_cell_index} -1 ){
        my $node = $self->{_cells}->[ $col_node_idx ];
        if(  $self->_get_cell_node_column($node) == $col_index ){
            return $self->_mk_cell( $col_node_idx );
        }
    }
    return;
}

sub _get_cell_node_column{
    my ($self, $cell_node) = @_;
    my $range = $cell_node->getAttribute( 'r' ) or return;
    my ( $book, $sheet, $row, $col ) = $self->{_sheet}->{_book}->parse_range( $range );
    #ignore book, sheet and row.
    return $col;
}

sub _mk_cell{
    my ($self, $col_index) = @_;
    my $cell_node = $self->{_cells}->[ $col_index ] or return;
    my $range = $cell_node->getAttribute( 'r' ) or return;
    # Create or re-use (for efficiency) a Cell object.
    my $cell = $self->{_cell};
    $cell->_init();
    $cell->{_range} = $range;
    #ignore book, sheet
    ( undef, undef, $cell->{_row}, $cell->{_col} ) = $self->{_sheet}->{_book}->parse_range( $range );
    my $type = $cell_node->getAttribute( 't' );
    $cell->{_type} = $type || '';
    # Read the cell <c> child nodes.
    for my $child_node ( $cell_node->childNodes() ) {
        my $node_name = $child_node->nodeName();
        if ( $node_name eq 'v' ) {
            $cell->{_value}     = $child_node->textContent();
            $cell->{_has_value} = 1;
        }

        if ( $node_name eq 'is' ) {
            $cell->{_value}     = $child_node->textContent();
            $cell->{_has_value} = 1;
        }
        elsif ( $node_name eq 'f' ) {
            $cell->{_formula}     = $child_node->textContent();
            $cell->{_has_formula} = 1;
        }
    }
    return $cell;
}

###############################################################################
#
# clone()
#
# Clone the current Row object.
#
sub clone {
	
		my $self = shift;
		
		my $clone = bless { %$self }, ref($self);
		
		return $clone;
}

###############################################################################
#
# values()
#
# Return an array of values for a row. The range is from the first cell up
# to the last cell. Returns '' for empty cells.
#
sub values {

    my $self = shift;
    my @values;

    # The row values are cached to allow multiple calls. Return cached values
    # if present.
    if ( defined $self->{_values} ) {
        return @{ $self->{_values} };
    }

    # Other wise read the values for the cells in the row.

    # Store any cell values that exist.
    while ( my $cell = $self->next_cell() ) {
        my $col   = $cell->col();
        my $value = $cell->value();
        $values[$col] = $value;
    }

    # Convert any undef values to an empty string.
    for my $value ( @values ) {
        $value = '' if !defined $value;
    }

    # Store the values to allow multiple calls return the same data.
    $self->{_values} = \@values;

    return @values;
}


###############################################################################
#
# row_number()
#
# Return the row number, zero-indexed.
#
sub row_number {

    my $self = shift;

    return $self->{_row_number};
}


###############################################################################
#
# previous_number()
#
# Return the zero-indexed row number of the previously found row. Returns -1
# if there was no previous number.
#
sub previous_number {

    my $self = shift;

    return $self->{_previous_row_number};
}


#
# Internal methods.
#


1;


__END__

=pod

=head1 NAME

Row - A class for reading Excel XLSX rows.

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
