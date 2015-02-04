package Excel::Reader::XLSX::Package::Styles;
use 5.008002;
use strict;
use warnings;
use Exporter;
use Carp;
use XML::Simple::Tiny qw(parsestring);
$XML::Simple::Tiny::TAGS  = 0;  #don't generate tag attributes.
$XML::Simple::Tiny::NAMES = 0;  #don't name a node by it's name attribute.

use XML::LibXML::Reader qw(:types);
use Excel::Reader::XLSX::Package::XMLreader;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';

our $FULL_DEPTH  = 1;
our $RICH_STRING = 1;


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    
    my $self  = bless Excel::Reader::XLSX::Package::XMLreader->new(), $class;
    
    return $self;
}

sub get_style{
    
    my $self    = shift;
    my $styleId = shift;
    
    my $styles  = {};
    return $styles unless $styleId;
    
    #retrive style and merge sub style into one parent hash
    my $parent = $self->get( cellxfs => $styleId );
    if($parent){
        if(exists $parent->{xfId}){
            my $cellStyle = $self->getByTag( cellStyle => xfId => $parent->{xfId} );
            $styles = $self->getBuiltin( $cellStyle->{xfId}, $cellStyle->{builtinId} );
            #copy buitins specifics
            $styles->{name}      = $cellStyle->{name};
            $styles->{builtinId} = $cellStyle->{builtinId};
        }
        $self->merge_styles( $styles, $parent );
    }
    
    return $styles;
}

sub get{
    
    my $self = shift;
    my $type = shift;
    my $id   = shift;
    
    my $hash = $self->{lc "_$type"}[$id];
    
    return $hash;
}

sub getByTag{
    
    my $self = shift;
    my $type = shift;
    my $tag  = shift;
    my $id   = shift;
    
    foreach my $hash( @{$self->{lc "_$type"}} ){
        return $hash 
            if exists($hash->{$tag}) and $hash->{$tag} eq $id;
    }
    
    return {};
}

sub getBuiltin{
    
    my $self        = shift;
    my $xfId        = shift;
    my $builtInId   = shift;
    
    #Builtins styles are defined in presetCellStyles.xml from [ECMA-376, 4th ed, part 1] zip archives
    
    #but if used in the workbook its in cellStyleXfs [ECMA-376, 4th ed, part 1 - page 1763]
    my $parent = $self->get( cellstylexfs => $xfId );
    # generate a new hash with all sub styles resolved on need / take in account apply* tags)
    my $styles = {};
    $styles    = $self->merge_styles( $styles, $parent );
    
    return $styles;
}

# colors can contain indexedColors if palette was modified and mruColors if custom colors was selected.

sub merge_styles{
    my $self   = shift;
    my $styles = shift;
    my $parent = shift;
    
    #exception for ApplyNumberFormat and ApplyAlignment (which is a child elements)
    my %apply    = map{ my $k = $_; s/^apply//; lc $_ => $parent->{$k} } grep{/^apply/}            keys %$parent;
    my %id       = map{ my $k = $_; s/Id$//;    lc $_ => $parent->{$k} } grep{/Id$/ and !/^xfId/ } keys %$parent;
    my @children = grep{ !/^apply|Id$|^(?:_id)$/ } keys %$parent;
    
    STYLE:
    for my $type( keys %id ){
        next STYLE if exists $apply{$type} and $apply{$type}==0;
        $styles->{$type} = $self->get( $type => $id{$type} );
    }
    
    #copy children (use reference)
    $styles->{$_} = $parent->{$_} for @children;

    return $styles;
}

sub get_indexed_color{
    
    my $self  = shift;
    my $index = shift;
    
    #reference extracted from ECMA-376, Part 4, Section 3.8.26
    #adapted from https://epplus.codeplex.com/discussions/231557
    my @indexed_color = ( 
            "FF000000", "FFFFFFFF", "FFFF0000", "FF00FF00", "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF",
            "FF000000", "FFFFFFFF", "FFFF0000", "FF00FF00", "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF",
            "FF800000", "FF008000", "FF000080", "FF808000", "FF800080", "FF008080", "FFC0C0C0", "FF808080",           
            "FF9999FF", "FF993366", "FFFFFFCC", "FFCCFFFF", "FF660066", "FFFF8080", "FF0066CC", "FFCCCCFF",
            "FF000080", "FFFF00FF", "FFFFFF00", "FF00FFFF", "FF800080", "FF800000", "FF008080", "FF0000FF",
            "FF00CCFF", "FFCCFFFF", "FFCCFFCC", "FFFFFF99", "FF99CCFF", "FFFF99CC", "FFCC99FF", "FFFFCC99",
            "FF3366FF", "FF33CCCC", "FF99CC00", "FFFFCC00", "FFFF9900", "FFFF6600", "FF666699", "FF969696",
            "FF003366", "FF339966", "FF003300", "FF333300", "FF993300", "FF993366", "FF333399", "FF333333",
            "FF000000", # indexed="64"    System Foreground
            "FFFFFFFF", # indexed="65"    System Background
        );
    
    return $indexed_color[$index];
}

sub get_theme_color{
    
    my $self  = shift;
    my $theme = shift;

    #TODO: need to parse theme1.xml (or theme[x].xml)
    
    return "FF000000";
}

sub apply_tint_color{
    
    my $self  = shift;
    my $color = shift;
    my $tint  = shift;

    #TODO
=pod
The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and 
1.0 means 100% lighten. Also, 0.0 means no change. 
 
In loading the RGB value, it is converted to HLS where HLS values are (0..HLSMAX), where 
HLSMAX is currently 255. 
 
[Example:  
 
Here are some examples of how to apply tint to color: 
 
If (tint < 0) 
  Lum’ = Lum * (1.0 + tint) 
 
For example: Lum = 200; tint = -0.5; Darken 50% 
  Lum‘ = 200 * (0.5) => 100 
 
For example:  Lum = 200; tint = -1.0; Darken 100% (make black) 

Attributes  Description 
  Lum‘ = 200 * (1.0-1.0) => 0 
 
If (tint > 0) 
  Lum‘ = Lum * (1.0-tint) + (HLSMAX – HLSMAX * (1.0-tint)) 
 
For example: Lum = 100; tint = 0.75; Lighten 75% 
Lum‘      = 100 * (1-.75)  + (HLSMAX – HLSMAX*(1-.75)) 
                = 100 * .25 + (255 – 255 * .25) 
                = 25 + (255 – 63) = 25 + 192 = 217 
 
For example: Lum = 100; tint = 1.0; Lighten 100% (make white) 
Lum‘      = 100 * (1-1)  + (HLSMAX – HLSMAX*(1-1)) 
                = 100 * 0 + (255 – 255 * 0) 
                = 0 + (255 – 0) = 255 
 
end example] 
=cut
    
    return $color;
    
}

##############################################################################
#
# _read_node()
#
# read each node and children to build styles cache.
#
sub _read_node {

    my $self = shift;
    my $node = shift;
    
    return unless $node->nodeType == XML_READER_TYPE_ELEMENT;
    
    my $reader = $self->{_reader};
    my $tag    = $node->name;
    
    #TODO: known unhandled tags are: 
    #       - tableStyles, tableStyleElement
    #       - dxf, dxfs : non cells elements formatting.
    #       - extLst (Extension List)
    
    if($tag =~ /^(?:font|fill|border|cellStyle|xf|colors)$/) {
        my $key = lc "_".( $tag eq 'xf' ? $self->{_xfparent} : $tag );
        my $id  = $self->{_ids}{$key}++;
        #quick hack
        my $node_as_xml = $node->readOuterXml;
        $node_as_xml =~ s/xmlns="[^"]+"//g;
        my $elmt = parsestring($node_as_xml)->{$tag};
        $elmt->{ _id } = $id;        
        push @{$self->{$key}}, $elmt;
    }elsif($tag =~ /^(?:cellXfs|cellStyleXfs)$/){
        $self->{_xfparent} = $tag;
    }
    
}

1;


__END__