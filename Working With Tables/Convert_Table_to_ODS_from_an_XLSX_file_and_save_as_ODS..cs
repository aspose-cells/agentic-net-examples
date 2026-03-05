using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file containing the table
        string sourcePath = "input.xlsx";

        // Desired path for the resulting ODS file
        string destPath = "output.ods";

        // Convert the XLSX workbook (including its tables) to ODS format
        // This uses the Aspose.Cells ConversionUtility as defined in the provided rule set.
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion from XLSX to ODS completed successfully.");
    }
}