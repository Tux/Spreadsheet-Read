requires   "Carp";
requires   "Data::Dumper";
requires   "Data::Peek";
requires   "Encode";
requires   "Exporter";
requires   "File::Temp"               => "0.22";  # ignore : CVE-2011-4116
requires   "List::Util";

recommends "Data::Dumper"             => "2.184";
recommends "Data::Peek"               => "0.52";
recommends "Encode"                   => "3.19";
recommends "File::Temp"               => "0.2311";
recommends "IO::Scalar";

on "configure" => sub {
    requires   "ExtUtils::MakeMaker";

    recommends "ExtUtils::MakeMaker"      => "7.22";

    suggests   "ExtUtils::MakeMaker"      => "7.70";
    };

on "test" => sub {
    requires   "Test::Harness";
    requires   "Test::More"               => "0.88";
    requires   "Test::NoWarnings";

    recommends "Test::More"               => "1.302195";
    };

feature "opt_csv", "Provides parsing of CSV streams" => sub {
    requires   "Text::CSV_XS"             => "0.71";

    recommends "Text::CSV"                => "2.02";
    recommends "Text::CSV_PP"             => "2.02";
    recommends "Text::CSV_XS"             => "1.50";
    };

feature "opt_gnumeric", "Provides parsing of Gnumeric spreadsheets" => sub {
    requires   "Spreadsheet::ReadGnumeric" => "0.2";

    recommends "Spreadsheet::ReadGnumeric" => "0.3";
    };

feature "opt_ods", "Provides parsing of OpenOffice spreadsheets" => sub {
    requires   "Spreadsheet::ParseODS"    => "0.26";

    recommends "Spreadsheet::ParseODS"    => "0.37";
    };

feature "opt_sxc", "Provides parsing of OpenOffice spreadsheets old style" => sub {
    requires   "Spreadsheet::ReadSXC"     => "0.26";

    recommends "Spreadsheet::ReadSXC"     => "0.37";
    };

feature "opt_tools", "Spreadsheet tools" => sub {
    recommends "Tk"                       => "804.036";
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
