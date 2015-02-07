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
my $json_filename = 't/regression/json_files/read_formula01.json';
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

SHEET:
for my $worksheet ( $workbook->worksheets() ) {

    my $sheetname = $worksheet->name();
    $got->{$sheetname} = [];
	ROW:
    while ( my $row = $worksheet->next_row() ) {
		CELL:
        while ( my $cell = $row->next_cell() ) {

            my $range   = $cell->range();
            my $formula = $cell->formula();
			next CELL unless defined $formula;
            push @{ $got->{$sheetname} },
              { range => $range, formula => $formula };
        }
    }
}


# Test the results.
my $test = _is_deep_diff( $got, $expected, $caption );

use constant DEBUG => 0;
if(DEBUG){
    unless($test){   #output wanted json
        use JSON;
        diag to_json($got, { utf8 => 1, pretty => 1 });
    }
}