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
my $json_filename = 't/regression/json_files/row/get_cell02.json';
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
#access values column by column.
foreach my $range ( [ 0,0 ], [ 2,0 ], [ 3,0 ], [ 4,0 ], [ 2,1 ], [0,2], [2,2], [4,2] ){
	my $row = $worksheet->get_row( $range->[0] ) 
					or next;
	my $cell = $row->get_cell($range->[1]) 
					or next;
	push @$got, $cell->value;
}

# Test the results.
_is_deep_diff( $got, $expected, $caption );
