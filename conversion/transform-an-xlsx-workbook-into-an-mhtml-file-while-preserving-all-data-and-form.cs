using System;
using Aspose.Cells.Utility;

class ConvertToMhtml
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired output path with .mht extension (MHTML format)
        string destPath = "output.mht";

        // Convert the Excel file to MHTML while preserving all data and formatting
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine($"Workbook successfully converted to MHTML: {destPath}");
    }
}