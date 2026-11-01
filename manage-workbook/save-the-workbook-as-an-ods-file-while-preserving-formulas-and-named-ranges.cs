// Title: Convert an Excel workbook to ODS format while preserving formulas and named ranges with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to save XLSX as ODS keeping formulas | preserve named ranges when converting Excel to OpenDocument Spreadsheet using Aspose.Cells | OdsSaveOptions example for retaining formulas in .NET | export workbook to .ods file with formula support in C# | Aspose.Cells convert Excel to ODS without losing named ranges
// Tags: xlsx to ods conversion with formula retention | named range preservation Aspose.Cells | OdsSaveOptions usage C# | export workbook to OpenDocument Spreadsheet | Aspose.Cells ODS format support | preserve formulas during ODS export

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX workbook, creates an OdsSaveOptions object (which by default keeps formulas and named ranges), and saves the workbook as an ODS file using Aspose.Cells for .NET, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\Path\To\Your\InputFile.xlsx";
        string outputPath = @"C:\Path\To\Your\OutputFile.ods";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // ODS save options – default behavior preserves formulas and named ranges
            OdsSaveOptions odsOptions = new OdsSaveOptions();

            // Save the workbook as an ODS file
            workbook.Save(outputPath, odsOptions);

            Console.WriteLine($"Workbook successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
