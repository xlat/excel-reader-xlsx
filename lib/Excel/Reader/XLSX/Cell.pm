package Excel::Reader::XLSX::Cell;

###############################################################################
#
# Cell - A class for reading the Excel XLSX cells.
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
use Excel::Reader::XLSX::Package::XMLreader;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $self  = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_sheet} = shift;
    $self->{_shared_strings} = shift;
    $self->{_value}          = '';

    bless $self, $class;

    return $self;
}


###############################################################################
#
# _init()
#
# Initialise a Cell object.
#
sub _init {

    my $self = shift;

    $self->{_value}            = '';
    $self->{_converted_string} = 0;
    $self->{_has_formula}      = 0;
    $self->{_formula}          = undef;
    $self->{_styles}           = undef;
    $self->{_styleId}          = undef;
    
}


###############################################################################
#
# value()
#
# Return the cell value.
#
sub value {

    my $self = shift;

    # If the cell type is a shared string convert the value index to a string.
    if ( $self->{_type} eq 's' && !$self->{_converted_string} ) {
        $self->{_value} =
          $self->{_shared_strings}->_get_string( $self->{_value} );

        # State variable so that multiple calls to value() don't need lookups.
        $self->{_converted_string} = 1;
    }


    return $self->{_value};
}


###############################################################################
#
# row()
#
# Return the cell row number, zero-indexed.
#
sub row {

    my $self = shift;

    return $self->{_row};
}


###############################################################################
#
# col()
#
# Return the cell column number, zero indexed.
#
sub col {

    my $self = shift;

    return $self->{_col};
}

###############################################################################
#
# range()
#
# Return the range of the current cell.
#
sub range {

    my $self = shift;

    return $self->{_range};
}

###############################################################################
#
# clone()
#
# return a clone of the current Cell object.
#
sub clone {

    my $self = shift;
    
    my $clone = bless { %$self }, ref($self);
    
    return $clone;
}


sub formula{
    
    my $self = shift;
    
    return $self->{_formula};
}


sub has_formula{
    
    my $self = shift;
    
    return $self->{_has_formula};
}

sub get_hyperlink{
    
    my $self = shift;
    
    if(my $formula = $self->formula){
        if($formula =~ /HYPERLINK\("([^"]+)","([^"]*)"\)/i){
            my ($range, $display) = ($1, $2);
            return { display => $display, location => $range };
        }
    }
    return undef;
}

sub styles{
    
    my $self = shift;

    #FIXME: take in account row->{_styleId} as a parent style to inherit from (if any).
    $self->{_styles} //= $self->{_sheet}{_styles}->get_style($self->{_styleId});
    
    return $self->{_styles};
}

#styles specific methods
sub color{

    my $self = shift;
    
    my $colors = $self->styles->{font}{color};

    return $self->{_sheet}{_styles}->get_color_as_rgb( $colors );
}

sub bgcolor{

    my $self = shift;
    
    #should be a patternType="solid"
    my $colors = $self->styles->{fill}{patternFill}{fgColor};

    return $self->{_sheet}{_styles}->get_color_as_rgb( $colors );
}

sub is_bold{
    
    my $self = shift;
    
    return exists $self->styles->{font}{b};
}

sub is_italic{
    
    my $self = shift;
    
    return exists $self->styles->{font}{i};
}

sub is_underline{
    
    my $self = shift;
    
    return exists $self->styles->{font}{u};
}

sub is_striketrough{
    
    my $self = shift;
    
    return exists $self->styles->{font}{strike};
}


sub is_superscript{
    
    my $self = shift;
    
	return exists $self->styles->{font}{vertAlign}
				? $self->styles->{font}{vertAlign} eq 'superscript'
				: 0
				;
}

sub is_subscript{
    
    my $self = shift;
    
	return exists $self->styles->{font}{vertAlign}
				? $self->styles->{font}{vertAlign} eq 'subscript'
				: 0
				;
}

sub is_merged{
	
	my $self = shift;
	
	return $self->{_sheet}->is_merged_rowcol($self->row, $self->col);
}

1;


__END__

=pod

=head1 NAME

Cell - A class for reading the Excel XLSX cells.

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
