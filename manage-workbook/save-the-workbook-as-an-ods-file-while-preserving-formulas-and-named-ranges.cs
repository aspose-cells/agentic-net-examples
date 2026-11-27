// Title: Convert an Excel workbook to ODS format while keeping formulas and named ranges using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert xlsx to ods preserving formulas | How to retain named ranges when saving workbook as ODS with Aspose.Cells | Saving Excel file as OpenDocument Spreadsheet using OdsSaveOptions in .NET | Aspose.Cells ODS export keep formulas and defined names in C# | Sample code to export .xlsx to .ods with formulas intact
// Tags: Aspose.Cells OdsSaveOptions configuration | save workbook as OpenDocument Spreadsheet in C# | retain formulas during ODS export | keep defined names when converting to ODS | C# convert XLSX to ODS using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the source XLSX file, loads it into an Aspose.Cells Workbook, configures OdsSaveOptions with SaveFormat.Ods, and saves the workbook as an ODS file, preserving both formulas and named ranges while handling any runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.ods";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure ODS save options (use the updated SaveFormat enum)
                OdsSaveOptions odsOptions = new OdsSaveOptions(SaveFormat.Ods);

                // Save the workbook as an ODS file
                workbook.Save(outputPath, odsOptions);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
