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
use XML::LibXML::Reader qw(:types);

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
    $self->{_styles}              = shift;
    $self->{_name}                = shift;
    $self->{_index}               = shift;
    my $state 						= shift;
	
    $self->{_previous_row_number} = -1;
    $self->{row_cache} = [] if $USE_CACHE;
	$self->{_properties} = { visibility => $state }; 
	$self->{_views} = [];
	$self->{_colsprops} = [];
	$self->{_mergedcells} = [];

    bless $self, $class;
	
    return $self;
}

##############################################################################
#
# _read_node()
#
# Callback function to read the nodes of the Worksheet.xml file.
#
sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only process the start elements.
    return unless $node->nodeType() == XML_READER_TYPE_ELEMENT;

    if ( $node->name eq 'tabColor' ) {
		#may have theme, tint and index?
		$self->{_properties}{tabcolor} = {
				rgb => $node->getAttribute( 'rgb' ),
			};
    }
	
    if ( $node->name eq 'dimension' ) {
		$self->{_properties}{dimension} = $node->getAttribute('ref');
	}
	
    if ( $node->name eq 'selection' ) {
		#and sqref attr?
		$self->{_properties}{selection} = $node->getAttribute('activeCell');
	}
	
    if ( $node->name eq 'sheetFormatPr' ) {
		#and sqref attr?
		$self->{_properties}{col_width} = $node->getAttribute('baseColWidth');
		$self->{_properties}{row_height} = $node->getAttribute('defaultRowHeight');
	}
	
    if ( $node->name eq 'col' ) {
		my $min = $node->getAttribute('min');
		my $max = $node->getAttribute('max');
		my $ref = {
				width => $node->getAttribute('width'),
				hidden => $node->getAttribute('hidden'),
				custom_width => $node->getAttribute('customWidth'),
			};
		$self->{_colsprops}[$_] = $ref for ($min-1..$max-1);
	}
	
    if ( $node->name eq 'mergeCell' ) {
		my $refs = $node->getAttribute('ref');
		my ($from, $to) = split /:/, $refs;
		my ($from_row, $from_col) = Excel::Reader::XLSX::Workbook::_range_to_rowcol($from);
		my ($to_row, $to_col) = Excel::Reader::XLSX::Workbook::_range_to_rowcol($to);
		($from_row, $to_row) = sort { $a <=> $b } ($from_row, $to_row);
		($from_col, $to_col) = sort { $a <=> $b } ($from_col, $to_col);
		push @{$self->{_mergedcells}}, sub{
			my ($row, $col) = (shift, shift);
			return $row>=$from_row && $row<=$to_row 
				&& $col >=$from_col && $col <=$from_col;
		};
	}
	
    if ( $node->name eq 'pageMargins' ) {
		$self->{_properties}{margins} = {
			left => $node->getAttribute('left'),
			right => $node->getAttribute('right'),
			top => $node->getAttribute('top'),
			bottom => $node->getAttribute('bottom'),
			header => $node->getAttribute('header'),
			footer => $node->getAttribute('footer'),
		};
	}	
	
    if ( $node->name eq 'setup' ) {
		$self->{_properties}{setup} = {
			paper_size =>  $node->getAttribute('paperSize'),
			orientation =>  $node->getAttribute('orientation'),
			#~ rId1 =>  $node->getAttribute('r:id'),
		};
	}
    
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
    my $row_style  = $row_reader->getAttribute( 's' );

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
    $row->_init( $row_number, $row_style );
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

#~ sub find_shared_formula{
	#~ my $self = shift;
	#~ my $si = shift;
	#~ (my $clone = $sheet->clone)->rewind;
	#~ ROW:
	#~ while(my $row = $clone->net_row){
		#~ CELL:
		#~ while(my $cell = $row->next_cell){
			#~ last ROW if defined $self->{_sheet}{_sharedformula}[$si];
		#~ }
	#~ }
	#~ return $self->{_sheet}{_sharedformula}[$si]; # or die "shared formua without source (si=$si)!";
#~ }

sub transpose_shared_formula{
	
	my $self = shift;
	my $si = shift;
	my $cell = shift;
	
	my $shared_formula = $self->{_sharedformula}[$si];
	my $formula = $shared_formula->{f};
	my ($src_row, $src_col) = Excel::Reader::XLSX::Workbook::_range_to_rowcol($shared_formula->{src});
	my $d_row	= $cell->row - $src_row;
	my $d_col	= $cell->col - $src_col;
	
	#find all CELL reference that are translatable in formula.
	#avoid to match function like LOG1( ) as a cell reference.
	#avoid text string such as in: =B1 & "A1"; to be patched.
	#	To find double-quoted strings in formula so we can ignore match inside them.
	#	Build a new string where non string chars are '0' and string chars are all '1' so 
	#	we could quickly know if a match is inside a string.
	my $ignore_match_map = join '', 
										map{ (/^"/o ? '1' : '0') x length } 
										split /("[^"]*")/o, $formula;
	my @tr;
	CELL_REF:
	while($formula =~ /\b([A-Z]+)(\$?)(\d+)\b(?![(])/go){
		my ($c, $rlock, $r, $i, $len) = ($1, $2, $3, $-[0], $+[0] -$-[0]);
		my $clock = $i ? substr($formula, $i - 1,1) eq '$' : 0;
		next CELL_REF if ($rlock and $clock) 
								or substr($ignore_match_map,$i,1);
		my ($row, $col) = Excel::Reader::XLSX::Workbook::_range_to_rowcol( $c.$r );
		$row += $d_row unless $rlock;
		$col += $d_col unless $clock;
		my $range = Excel::Reader::XLSX::Workbook::_rowcol_to_range( $row, $col, $rlock, $clock );
		#patch range, in a reverse order
		unshift @tr, sub{ substr($_[0], $i, $len)=$range };
	}
	$_->($formula) for @tr;
	
	return $formula;
}

sub resolve_external_workbook{
	my $self = shift;
	my $formula = shift;
	
	my $ignore_match_map = join '', 
										map{ (/^"/ ? '1' : '0') x length } 
										split /("[^"]*")/, $formula;

	my $rx_ext_wb = qr/
		(?<quote>'\[)(?<wb>\d+)\](?<ws>(?:[^']||'')+)'!
	|	\[(?<wb>\d+)\](?<ws>[^'!]*)!
		/x;
	my @tr;
	EXTERNAL_WB:
	while($formula =~ /$rx_ext_wb/g){
		my ($quote, $extId, $sheet, $i, $len) = ($+{quote}, $+{wb}, $+{ws}, $-[0], $+[0] - $-[0]);
		next EXTERNAL_WB if substr($ignore_match_map,$i, 1);
		$sheet//='';#names ref doesn't have sheet name, so avoid undef.
		my $wbref = $self->{_book}->get_external_target($extId - 1);
		unless($wbref =~ s{^(.*)/([^/]+)$}{'$1/[$2]$sheet'!}){
			if($quote){
				$wbref = "'[$wbref]$sheet'!";
			}
			else{
				$wbref = "[$wbref]$sheet!";
			}
		}
		$wbref =~ tr{/}{\\};
		unshift @tr, sub{ substr($_[0], $i, $len)=$wbref };
	}
	$_->($formula) for @tr;
	
=pod

	[1]Sheet1!$C$3 							=> [data01.xlsx]Sheet1!$C$3
	'[2]Sheet With Spaces'!$A$1			=> 'externals\[external01.xlsx]Sheet1'!$A$1
	'[2]Sheet''s "nam({e})!"'!$A$1	=> 'externals\[external01.xlsx]Sheet''s "nam({e})!"'!$A$1
	
	[2]!EXTNAME1		=> 'externals\[external01.xlsx]Sheet'!$A$1
	[2]!NAMED			=> 'externals\[external01.xlsx]Sheet With Spaces'!$A$2

=cut	
	return $formula;
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
    $self->_parse_file( $filename );
    $self->SUPER::rewind;
	
}

sub property{
	
	my $self = shift;
	my $propname = shift;
	
	return $self->{_properties}{$propname} 
		if exists $self->{_properties}{$propname};
	return undef;
}

sub properties{
	my $self = shift;
	
	return keys %{$self->{_properties}};
}

sub col_width{
	my $self = shift;
	my $col = shift;
	
	my $width = $self->{_colsprops}[$col] 
						// $self->property('col_width');
}

sub is_merged_rowcol{
	my $self = shift;
	my $row = shift;
	my $col  = shift;
	
	for(@{$self->{_mergedcells}}){
		my $ismerged = $_->($row, $col);
		return $ismerged if $ismerged;
	}
	return 0;
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
