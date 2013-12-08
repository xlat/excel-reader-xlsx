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
my $json_filename = 't/regression/json_files/workbook/parse_range.json';
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
$workbook->{_names} = { 'NAMED' => 'Sheet1!D3' };
my @ranges = ( 
 'A1', 									
 'B2', 									
 'AA1', 									
 'AB1', 									
 'NAMED', 								
 '$A1', 									
 'A$1', 									
 '$A$1', 								
 'Sheet1!A1', 						
 '\'Sheet1\'!A1', 						
 'Sheet1!$A$1', 						
 '\'Sheet1\'!$A$1', 					
 'data01.xlsx!NAMED', 				
 '[data01.xlsx]Sheet1!NAMED',	
 '[data01.xlsx]Sheet1!$C$2',	
 'ZZZZZ',
	);
for my $range ( @ranges ) {

    my ($book, $sheet, $row, $col) = $workbook->parse_range( $range );
    push @$got, [ $range, $book, $sheet, $row, $col ];

}

# Test the results.
_is_deep_diff( $got, $expected, $caption );
