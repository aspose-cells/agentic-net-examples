// Title: Convert an Excel .xlsx workbook to OpenDocument Spreadsheet (.ods) in C# while retaining formulas and named ranges using Aspose.Cells
// AI Prompts: Generate C# code that loads an existing .xlsx file with Aspose.Cells, configures OdsSaveOptions to keep formulas and named ranges, and saves it as a .ods file. | Show how to verify that formulas and named ranges remain intact after converting an Excel workbook to ODS using Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells C# export to ODS retaining calculations | keep defined names when converting Excel to OpenDocument format in .NET | C# OdsSaveOptions save workbook as .ods without losing spreadsheet logic | how to convert .xlsx to .ods with Aspose.Cells preserving cell references | export Excel workbook to ODS while maintaining all formulas in C#
// Tags: C# OdsSaveOptions retain spreadsheet logic | Aspose.Cells ODS export with defined names | Excel to OpenDocument conversion using Aspose.Cells | save workbook as .ods preserving calculations | load .xlsx and export .ods Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, creates default OdsSaveOptions (which automatically keep formulas and named ranges), and saves the workbook as output.ods, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Ensure the source file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // ODS save options (default behavior preserves formulas and named ranges)
            var odsOptions = new OdsSaveOptions();

            // Save the workbook as ODS
            workbook.Save(outputPath, odsOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
