#!/usr/bin/perl -w

use strict;
use warnings;
use Spreadsheet::ParseExcel;
use Getopt::ArgParse;
use Spreadsheet::ParseXLSX;



my $ap = Getopt::ArgParse->new_parser(
        prog => 'rd_xls',
        help => 'usage: rd_xls -d <delimiter> -x <excel file> [-c,--colors -b,--borders -f,--fonts]',
        description => 'parses excel files',
        epilog => 'by dan mercurio'
    );
$ap->add_arg('--delimiter','-d',required => 1);
$ap->add_arg('--excel_file','-x',required => 1);
$ap->add_arg('--borders','-b',type => 'Bool',required => 0);
$ap->add_arg('--colors','-c',type => 'Bool',required => 0);
$ap->add_arg('--font','-f',type => 'Bool', required => 0);

my $ns = $ap->parse_args();

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

if (substr($ns->excel_file,-4) eq 'xlsx') {

    #print "\nxlsx file detected\n";
    $excelparser = Spreadsheet::ParseXLSX->new();
    $workbook = $excelparser->parse($ns->excel_file);

    if ( !defined $workbook ) {
        die $excelparser->error(), ".\n";
    }


} else {

    $excelparser = Spreadsheet::ParseExcel->new();
    $workbook = $excelparser->parse($ns->excel_file);



    if ( !defined $workbook ) {
        die $excelparser->error(), ".\n";
    }

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
                if ($ns->borders) {
                    &printBorderType($cell);
                }
                if ($ns->colors) {
                    &printCellColor($cell);
                }
                if ($ns->font) {
                    &printFontInfo($cell);
                }
                print $dl;
        }
    }
}