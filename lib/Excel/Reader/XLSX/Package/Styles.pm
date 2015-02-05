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
    
    my $self        = shift;
    my $theme_index = shift;

    #TODO: need to parse theme1.xml (or theme[x].xml) 
    #               for theme/themeElements/clrScheme/*/sysClr,srgbClr
    #Currently, it comes from default values with default [windows] theme.
    my @themes = (
        "FF000000", "FFFFFFFF", "FF1F497D", "FFEEECE1", 
        "FF4F81BD", "FFC0504D", "FF9BBB59", "FF8064A2", 
        "FF4BACC6", "FFF79646", "FF0000FF", "FF800080", 
    );
        
    return $themes[$theme_index] // "FF000000";
}

sub get_color_as_rgb{
    
    my $self   = shift;
    my $colors = shift;
    my $color;
    
    $color = $colors->{rgb} if exists $colors->{rgb};
    $color = $self->get_indexed_color($colors->{indexed}) if exists $colors->{indexed};
    $color = $self->get_theme_color($colors->{theme}) if exists $colors->{theme};
    $color = apply_tint_to_argb($colors->{tint}, $color) if exists $colors->{tint};
    
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

################ 
# TODO: move the coloration functions into a specific module like Excel::Reader::XLSX::Colors or Color::HSL
use List::Util qw( min max );
use Math::Round qw( round );

sub rgb_to_hex{
    my @rgb  = (shift, shift, shift);
    my %opts = @_;
    unshift @rgb, $opts{argb} if exists $opts{argb};
    my $hex;
    $hex = '#' if $opts{html};
    $hex .= sprintf('%02X', $_) for @rgb;
    return $hex;
}
sub hex_to_rgb{
    my $argb = shift;
    
    if($argb =~ /^(?:#|..|&[hH]|0[xX])?(..)(..)(..)$/){
        return map{ hex_to_dec($_) } ($1, $2, $3);
    }
    
    return (0, 0, 0);
}

sub hex_to_dec{
    my $hex     = uc shift;
    
    my $v       = 0;
    my @hdigits = split //, $hex;
    for(@hdigits){
        $v *= 16;
        $v += index('0123456789ABCDEF',$_);
    }
    
    return $v;
}
sub rgb_to_hsl{
    my ($R, $G, $B) = map { $_ / 255 } @_;
    
    my ($H, $S, $L) = (0, 0, 0);
    my $Cmax = max( $R, $G, $B );
    my $Cmin = min( $R, $G, $B );
    my $delta= $Cmax - $Cmin;
    $L = ($Cmax + $Cmin) / 2;
    if($delta){
        if($Cmax == $R){
            $H = 60 * ((($G - $B)/$delta) % 6);
        }
        elsif($Cmax == $G){
            $H = 60 * (($B - $R)/$delta + 2);
        }
        else{
            $H = 60 * (($R - $G)/$delta + 4);
        }
        my $l = 1/3*($R + $G + $B);
        $S = 1 - $Cmin / $l;
    }
    $H = round( $H/360 * 255 );
    $S = round( $S * 255 );
    $L = round( $L * 255 );
    
    return ($H, $S, $L);
}
sub hsl_to_rgb{
    my ($H, $S, $L) = @_;
    my ($R, $G, $B) = (0, 0, 0);
    
    $H = $H * 360 / 255;
    $S = $S / 255;
    $L = $L / 255;

    my $C = (1 - abs( 2* $L - 1) ) * $S;
    $H /= 60;
    my $X = $C * ( 1 - abs( ( $H % 2 ) - 1 ) );
    my $m = $L - $C/2;
    # ($R, $G, $B) =  (0, 0, 0) if $H is undefined !
    if($H >=0 and $H < 1){
        ($R, $G, $B) = ( $C, $X, 0);
    }
    elsif($H >=1 and $H <2){
        ($R, $G, $B) = ( $X, $C, 0);
    }
    elsif($H >=2 and $H <3){
        ($R, $G, $B) = ( 0, $C, $X);
    }
    elsif($H >=3 and $H <4){
        ($R, $G, $B) = ( 0, $X, $C);
    }
    elsif($H >=4 and $H <5){
        ($R, $G, $B) = ( $X, 0, $C);
    }
    else{
        ($R, $G, $B) = ( $C, 0, $X);
    }

    $_ += $m for ($R, $G, $B);
    $_ = round($_ * 255) for ($R, $G, $B);
    
    return ($R, $G, $B);
}
sub apply_tint_to_hsl{
    my $tint = 1 + shift;
    my ($H, $S, $L)  = @_;
    
    $L = round(
            $tint<0
            ?   $L*(1+$tint) 
            :   $L*(1-$tint)+(255-255*(1-$tint))
        );
        
    return ($H, $S, $L);
}
sub apply_tint_to_rgb{
    my $tint = shift;
    my @rgb = @_;
    my @hsl  = rgb_to_hsl(@rgb);
    @hsl     = apply_tint_to_hsl( $tint, @hsl );
    @rgb     = hsl_to_rgb( @hsl );
    return @rgb;
}
sub apply_tint_to_argb{
    my $tint = shift;
    my $argb = shift;
    
    my @rgb  = hex_to_rgb($argb);
    @rgb     = apply_tint_to_rgb( $tint, @rgb);
    $argb    = rgb_to_hex(@rgb, argb => 255);
    
    return $argb;
}

1;


__END__