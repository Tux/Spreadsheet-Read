requires   "Carp";
requires   "Data::Dumper";
requires   "Data::Peek";
requires   "Encode";
requires   "Exporter";
requires   "File::Temp"               => "0.22";
requires   "List::Util";

recommends "Data::Peek"               => "0.49";
recommends "File::Temp"               => "0.2309";
recommends "IO::Scalar";

on "configure" => sub {
    requires   "ExtUtils::MakeMaker";
    };

on "test" => sub {
    requires   "Test::Harness";
    requires   "Test::More"               => "0.88";
    requires   "Test::NoWarnings";

    recommends "Test::More"               => "1.302175";
    };

feature "opt_csv", "Provides parsing of CSV streams" => sub {
    requires   "Text::CSV_XS"             => "0.71";
    
    recommends "Text::CSV"                => "2.00";
    recommends "Text::CSV_PP"             => "2.00";
    recommends "Text::CSV_XS"             => "1.44";
    };

feature "opt_ods", "Provides parsing of OpenOffice spreadsheets" => sub {
    requires   "Spreadsheet::ParseODS"    => "0.25";
    };

feature "opt_sxc", "Provides parsing of OpenOffice spreadsheets old style" => sub {
    requires   "Spreadsheet::ReadSXC"     => "0.24";
    
    recommends "Spreadsheet::ReadSXC"     => "0.25";
    };

feature "opt_tools", "Spreadsheet tools" => sub {
    recommends "Tk"                       => "804.035";
    recommends "Tk::NoteBook";
    recommends "Tk::TableMatrix::Spreadsheet";
    };

feature "opt_xls", "Provides parsing of Microsoft Excel files" => sub {
    requires   "Spreadsheet::ParseExcel"  => "0.34";
    requires   "Spreadsheet::ParseExcel::FmtDefault";
    
    recommends "Spreadsheet::ParseExcel"  => "0.65";
    };

feature "opt_xlsx", "Provides parsing of Microsoft Excel 2007 files" => sub {
    requires   "Spreadsheet::ParseExcel::FmtDefault";
    requires   "Spreadsheet::ParseXLSX"   => "0.24";
    
    recommends "Spreadsheet::ParseXLSX"   => "0.27";
    };
