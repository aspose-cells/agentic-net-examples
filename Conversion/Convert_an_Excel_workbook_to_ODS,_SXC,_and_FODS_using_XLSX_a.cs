using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Convert XLSX to OpenDocument Spreadsheet (ODS)
        string odsPath = "output.ods";
        ConversionUtility.Convert(sourcePath, odsPath);

        // Convert XLSX to StarOffice Calc Spreadsheet (SXC)
        string sxcPath = "output.sxc";
        ConversionUtility.Convert(sourcePath, sxcPath);

        // Convert XLSX to OpenDocument Flat XML Spreadsheet (FODS)
        string fodsPath = "output.fods";
        ConversionUtility.Convert(sourcePath, fodsPath);

        Console.WriteLine("All conversions completed successfully.");
    }
}