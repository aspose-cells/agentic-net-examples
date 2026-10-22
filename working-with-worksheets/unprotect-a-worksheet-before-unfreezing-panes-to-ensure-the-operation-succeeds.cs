// Title: How to unprotect an Excel worksheet and remove frozen panes with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that opens a workbook, calls Worksheet.Unprotect(), then clears any frozen panes with Worksheet.FreezePanes(0,0,0,0), and saves the file. | Show a try‑catch example that verifies the source file, creates missing output directories, and performs worksheet unprotection before unfreezing panes in Aspose.Cells. | Demonstrate how to programmatically reset pane freezing on a protected sheet after calling Unprotect() in a .NET application with Aspose.Cells.
// Common Searches: aspnet unprotect worksheet before using FreezePanes in Aspose.Cells | remove frozen panes from a protected Excel file using C# Aspose.Cells | how to reset frozen panes after worksheet protection with Aspose.Cells .NET | C# code to unprotect Excel sheet and clear pane freezing using Aspose.Cells library
// Tags: aspocells unprotect method usage | aspocells freezepanes reset example | c# modify worksheet protection aspocells | excel pane clearing with aspocells | aspocells workbook save after protection change

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that input.xlsx exists, loads it with Aspose.Cells, accesses the first worksheet, calls Unprotect() to lift protection, clears any frozen panes using FreezePanes(0,0,0,0), ensures the output directory is present, and saves the modified workbook to output.xlsx while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet to allow modifications
            worksheet.Unprotect();

            // Unfreeze any frozen panes (using FreezePanes method with zero parameters)
            worksheet.FreezePanes(0, 0, 0, 0);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
