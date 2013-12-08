###############################################################################
#
# Tests for Excel::Writer::XLSX.
#
# reverse('(c)'), February 2012, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_is_deep_diff _read_json);
use strict;
use warnings;
use Excel::Reader::XLSX;

use Test::More tests => 1;

###############################################################################
#
# Test setup.
#
my $json_filename = 't/regression/json_files/worksheet/get_range01.json';
my $json          = _read_json( $json_filename );
my $caption       = $json->{caption};
my $expected      = $json->{expected};
my $xlsx_file     = 't/regression/xlsx_files/' . $json->{xlsx_file};
my $got;


###############################################################################
#
# Test reading data from an Excel file.
#
use Excel::Reader::XLSX;

my $reader   = Excel::Reader::XLSX->new();
my $workbook = $reader->read_file( $xlsx_file );

my $worksheet = $workbook->worksheet(0);
#sequential access
foreach my $range ( "A1", "C1", "A3", "B3", "C3", "A4", "A5", "B5", "C5"){
	my $cell = $worksheet->get_range($range);
	push @$got, $cell->value 
		if ref $cell eq 'Excel::Reader::XLSX::Cell';
}

# Test the results.
_is_deep_diff( $got, $expected, $caption );
