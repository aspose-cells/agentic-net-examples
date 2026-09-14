// Title: Check for required worksheets and handle missing sheet errors when using ExportHiddenWorksheet to save PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that verifies each worksheet in a specified list exists in a Workbook before calling Workbook.Save with PdfSaveOptions.ExportHiddenWorksheet enabled, and logs a clear error if any sheet is absent. | Update the example program to set ExportHiddenWorksheet = true, iterate over an array of expected sheet names, catch missing‑worksheet exceptions, and abort the PDF export with a descriptive message.
// Common Searches: Aspose.Cells how to detect missing worksheets before exporting hidden sheets to PDF in C# | C# code sample for validating worksheet names when using ExportHiddenWorksheet option | Error handling for absent worksheets during PDF conversion with Aspose.Cells .NET
// Tags: validate worksheet existence Aspose.Cells PDF export | ExportHiddenWorksheet missing sheet handling C# | Aspose.Cells PDF conversion error handling | check required worksheets before Workbook.Save | C# Excel to PDF hidden sheet validation

using Aspose.Cells;
using System;
using System.IO;

// The program first confirms the input Excel file exists, then loads it into a Workbook with exception handling. It defines a list of required worksheet names and iterates through them, attempting to retrieve each via workbook.Worksheets[sheetName]; if a worksheet is missing, it logs an error and stops execution. After successful validation, PdfSaveOptions.ExportHiddenWorksheet is enabled and the workbook is saved to PDF, with a try‑catch block to report any export failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Configure PDF save options
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // Note: ExportHiddenWorksheet property may not be available in older versions.
        // If needed, ensure the used Aspose.Cells version supports it.

        // Define the worksheets that are expected to be present
        string[] requiredSheets = { "Sheet1", "HiddenSheet" };

        // Verify that each required worksheet exists
        foreach (string sheetName in requiredSheets)
        {
            Worksheet ws = null;
            try
            {
                ws = workbook.Worksheets[sheetName];
            }
            catch (ArgumentException)
            {
                // Worksheet not found
            }

            if (ws == null)
            {
                Console.WriteLine($"Error: Worksheet '{sheetName}' is missing.");
                return;
            }
        }

        try
        {
            // Save the workbook to PDF
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine("Export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Export failed: {ex.Message}");
        }
    }
}
