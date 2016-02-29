#!/usr/bin/perl -w

use strict;
use Spreadsheet::ParseExcel;
use Getopt::ArgParse;

my $ap = Getopt::ArgParse->new_parser(
        prog => 'rd_xls',
        help => 'usage: rd_xls <delimiter> <excel file> [-a b|o|c]',
        description => 'parses excel files',
        epilog => 'by dan mercurio'
    );

$ap->add_arg('delimiter', required => 1);
$ap->add_arg('excel_file', required => 1);
$ap->add_arg('-a', required => 0);


my $ns = $ap->parse_args();

my $printBorderOption = 0;
my $printCellColorOption = 0;
my $printFontOption = 0;

print "Delimiter: ".$ns->delimiter."\n";
print "File: ".$ns->excel_file."\n";
my @options = split("", $ns->a);

if ("b" ~~ @options) {
    $printBorderOption = 1;
}
if ("o" ~~ @options) {
    $printFontOption = 1;
}
if ("c" ~~ @options) {
    $printCellColorOption = 1;
}

#print border type subroutine
sub printBorderType {
    my $cell = $_[0];
    my $format = $cell->get_format();
    my @borderArray = $format->{BdrStyle}; 
            my $left = $borderArray[0][0];
            my $right = $borderArray[0][1];
            my $top = $borderArray[0][2];
            my $bottom = $borderArray[0][3];
            print " Borders: Left: " . $left;
            print ", Right: " . $right;
            print ", Top: " . $top;
            print ", Bottom: " . $bottom;
            print " ";

}
#print cell color subroutine
sub printCellColor {
    my $cell = $_[0];
    my $format = $cell->get_format();
    print " Color: ";
    if (!defined($format->{Fill}->[1])) {
        print "None"
    } else {
        print $format->{Fill}->[1];

    }
    print " ";
    #array location 1 is the FOREGROUND color
}

#print font info subroutine
sub printFontInfo {
    print " Font: ";
    my $cell = $_[0];
    my $format = $cell->get_format();
    print $format->{Font}->{Name};
    print " ";
}

my $workbook;
my $excelparser;

$excelparser = Spreadsheet::ParseExcel->new();
$workbook = $excelparser->parse($ns->excel_file);

if ( !defined $workbook ) {
    die $excelparser->error(), ".\n";
}

for my $worksheet ( $workbook->worksheets() ) {

    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();

    for my $row ( $row_min .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {
            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;

                my $dl;

                if ($ns->delimiter eq 'tb' or $ns->delimiter eq 'tab' or $ns->delimiter eq '\t') {
                    $dl = "\t";
                } else {
                    $dl = $ns->delimiter;
                }
                bless $cell, "Spreadsheet::ParseExcel::Cell";
                my $val = $cell->value();
                $val =~ s/$dl/(delimiter)/g;
                #$val =~ s/\W//g;
                $val =~ s/\n|\r/(character stripped)/g;
                print $val;
                if ($printBorderOption == 1) {
                    &printBorderType($cell);
                }
                if ($printCellColorOption == 1) {
                    &printCellColor($cell);
                }
                if ($printFontOption == 1) {
                    &printFontInfo($cell);
                }
                print $dl;
        }
    }
}
