package Excel::Reader::XLSX::Workbook;

###############################################################################
#
# Workbook - A class for reading the Excel XLSX workbook.xml file.
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
use Excel::Reader::XLSX::Worksheet;
use Excel::Reader::XLSX::Package::Relationships;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';


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

    my $class          = shift;
    my $package_dir    = shift;
    my $shared_strings = shift;
    my $styles         = shift;
    my %files          = @_;

    my $self = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_package_dir}          = $package_dir;
    $self->{_shared_strings}       = $shared_strings;
    $self->{_styles}               = $styles;
    $self->{_files}                = \%files;
    $self->{_worksheets}           = undef;
    $self->{_worksheet_properties} = [];
    $self->{_worksheet_indices}    = {};

    # Set the root dir for the workbook and worksheets. Usually 'xl/'.
    $self->{_workbook_root} = $self->{_files}->{_workbook};
    $self->{_workbook_root} =~ s/workbook.xml$//;

    bless $self, $class;

    $self->_set_relationships();

    return $self;
}


###############################################################################
#
# _set_relationships()
#
# Set up the Excel relationship links between package files and the
# internal ids.
#
sub _set_relationships {

    my $self     = shift;
    my $filename = shift;

    my $rels_file = Excel::Reader::XLSX::Package::Relationship->new();

    $rels_file->_parse_file(
        $self->{_package_dir} . $self->{_files}->{_workbook_rels} );

    my %rels = $rels_file->_get_relationships();
    $self->{_rels} = \%rels;
}


##############################################################################
#
# _read_node()
#
# Callback function to read the nodes of the Workbook.xml file.
#
sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only process the start elements.
    return unless $node->nodeType() == XML_READER_TYPE_ELEMENT;

    if ( $node->name eq 'sheet' ) {

        my	$name     = $node->getAttribute( 'name' );
        my	$sheet_id = $node->getAttribute( 'sheetId' );
        my	$rel_id   = $node->getAttribute( 'r:id' );
            $rel_id   =~ /(\d+)/;
        my	$index    = $1 - 1;
        # Use the package relationship data to convert the r:id to a filename.
        my	$filename = $self->{_rels}->{$rel_id}->{_target};

        # Store the properties to set up a Worksheet reader object.
        push @{ $self->{_worksheet_properties} },
          {
            _name     => $name,
            _sheet_id => $sheet_id,
            _index    => $index,
            _rel_id   => $rel_id,
            _filename => $filename,
          };
    }
    
    if( $node->name eq 'definedName' ) {
        # why $node->value( ) doesn't works?
        my $name = $node->getAttribute('name');
        my $value = $node->readInnerXml;
        $self->{_names}->{ uc $name } = $value;
        # other attributes :
        #    localSheetId="68" 
        #    hidden="1"
        #    and maybe others like comment
    }
    
}


###############################################################################
#
# worksheets()
#
# Return an array of Worksheet objects.
#
sub worksheets {

    my $self = shift;

    # Read the worksheet data if it hasn't already been read.
    if ( !defined $self->{_worksheets} ) {
        $self->_read_worksheets();
    }

    return @{ $self->{_worksheets} };
}


###############################################################################
#
# worksheet()
#
# Return a Worksheet object based on its sheetname or index. Unknown sheet-
# names or out of range indices return an undef object.
#
sub worksheet {

    my $self  = shift;
    my $index = shift;
    my $name  = $index;

    # Ensure some parameter was passed.
    return unless defined $index;

    # Read the worksheet data if it hasn't already been read.
    if ( !defined $self->{_worksheets} ) {
        $self->_read_worksheets();
    }

    # Convert a valid sheetname to an index.
    if ( exists $self->{_worksheet_indices}->{$name} ) {
        $index = $self->{_worksheet_indices}->{$name};
    }

    # Check if it is a valid index.
    return if $index !~ /^[-\d]+$/;

    return $self->{_worksheets}->[$index];
}



###############################################################################
#
# parse_range()
#
# Return book, sheet, row, column and range extracted from given $range.
# This method will resolve internal names but not interbook names (at least not yet) nor list of ranges.
sub parse_range{
    my ($self, $range) = @_;
    my ($book, $sheet, $row, $col) = (undef, undef, undef, undef);
    if($range =~ /^\[(?<book>[^\]]+)\](?<sheet>[^!]+)!(?<range>.*)/){
        $book = $+{book};
        $sheet = $+{sheet};
        $range = $+{range};
    }
    elsif($range =~ /^(?<book>[^.]+\.[^!]+)!(?<range>.+)/){
        $book = $+{book};
        $range = $+{range};
    }
resolve_names: 
    do{
        if(exists $self->{_names}->{uc $range}){
            #resolve name
            $range = $self->{_names}->{uc $range};
            #this new $range can contain an Sheet! prefix
            return if $range eq '#REF!';
        }
        if($range =~ /^(?<sheet>[^!]+)!(?<range>.+)/){
            $sheet = $+{sheet} if exists $+{sheet};
            $range = $+{range};
        }
    }while( $range =~ /!/ or exists $self->{_names}->{uc $range});
    $sheet =~ s/'//g if defined $sheet;
    my @refs;
    foreach $range ( split /[,;]/, $range ){
        if($range =~/^([^:]+):(.*)$/){
            my ($start, $end) = split ':', $range;
            my ($start_row, $start_col) = _range_to_rowcol( $start );
            my ($end_row,   $end_col)   = _range_to_rowcol( $end );
            foreach $row ($start_row .. $end_row){
                foreach $col ($start_col .. $end_col){
                    my $subrange = _rowcol_to_range( $row, $col );
                    push @refs, [$book, $sheet, $row, $col, $subrange];
                }
            }
        }
        else{
            ($row, $col) = _range_to_rowcol( $range );
            $range =~ s/\$//g;
            push @refs, [$book, $sheet, $row, $col, $range];
        }
    }
    return @refs;
}

###############################################################################
#
# Internal methods.
#
###############################################################################
###############################################################################
#
# _range_to_rowcol($range)
#
# Convert an Excel A1 style ref to a zero indexed row and column.
#
sub _range_to_rowcol {
    my $range = shift or return;
    $range =~s/\$//g;
    my ( $col, $row ) = split /(\d+)/, $range;
    return unless defined $row;
    $row--;

    my $length = length $col;

    if ( $length == 1 ) {
        $col = -65 + ord( $col );
    }
    elsif ( $length == 2 ) {
        my @chars = split //, $col;
        $col = -1729 + ord( $chars[1] ) + 26 * ord( $chars[0] );
    }
    else {
        my @chars = split //, $col;
        $col =
          -44_993 +
          ord( $chars[2] ) +
          26 * ord( $chars[1] ) +
          676 * ord( $chars[0] );
    }

    return $row, $col;
}

sub _rowcol_to_range {
    my ($row, $col, $rlock, $clock) = @_;
    my $range;
    $range = '$' if $clock;
    if( $col > 26 ){
        if($col > 701){
            $range = chr( int(int(($col / 26) / 26) + 64) ). 
                     chr( int(($col / 26) % 26 + 64) ). 
                     chr( int($col % 26 + 65) );
        }else{
            $range = chr( int($col / 26 + 64) ). 
                     chr( int($col % 26 + 65) );
        }
    }
    else{
        $range = chr( $col + 65 );
    }
    $range .= '$' if $rlock;    
    $range .= $row + 1;
    return $range;
}

###############################################################################
#
# _read_worksheets()
#
# Parse the workbook and set up the Worksheet objects.
#
sub _read_worksheets {

    my $self = shift;

    # Return if the worksheet data has already been read.
    return if defined $self->{_worksheets};

    # Iterate through the worksheet properties and set up a Worksheet object.
    for my $sheet ( @{ $self->{_worksheet_properties} } ) {

        # Create a new Worksheet reader.
        my $worksheet = Excel::Reader::XLSX::Worksheet->new(
            $self->{_shared_strings},
            $self->{_styles},
            $sheet->{_name},
            $sheet->{_index},
        );

        # Set up the file to read. We don't read data until it is required.
        $worksheet->_init( $self, $sheet );

        # Store the Worksheet reader objects.
        push @{ $self->{_worksheets} }, $worksheet;

        # Store the Worksheet index so it can be looked up by name.
        $self->{_worksheet_indices}->{ $sheet->{_name} } = $sheet->{_index};
    }
}



1;


__END__

=pod

=head1 NAME

Workbook - A class for reading the Excel XLSX workbook.xml file.

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
