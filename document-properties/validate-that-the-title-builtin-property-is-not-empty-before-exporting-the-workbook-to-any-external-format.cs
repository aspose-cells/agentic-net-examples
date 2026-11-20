// Title: Check that the Excel workbook's built‑in Title property is populated before saving to PDF with Aspose.Cells for .NET
// Common Searches: how to verify Excel built‑in Title property with Aspose.Cells before PDF export in C# | Aspose.Cells C# prevent saving workbook to PDF when Title document property is empty | validate workbook metadata Title using Aspose.Cells .NET API | C# code to throw error if Excel Title property is missing before converting to PDF with Aspose.Cells | ensure document property Title is set in Aspose.Cells before calling Save as PDF
// Tags: validate built-in Title property Aspose.Cells | prevent PDF export missing Title Aspose.Cells | Excel workbook metadata check C# | Aspose.Cells SaveFormat.Pdf property validation | C# workbook property validation before conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, confirms the file exists, handles password‑protected workbooks, retrieves the built‑in Title property, throws an InvalidOperationException if the Title is empty or whitespace, and then saves the workbook as a PDF, with comprehensive exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook (handle possible password protection)
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex) when (ex.Message.Contains("Please provide password"))
            {
                // If the workbook is password‑protected, inform the user and abort gracefully
                Console.WriteLine("The workbook is password‑protected. Provide a valid password to load it.");
                return;
            }

            // Retrieve the built‑in Title property
            string title = workbook.BuiltInDocumentProperties["Title"]?.ToString();

            // Validate that the Title is not empty or whitespace
            if (string.IsNullOrWhiteSpace(title))
                throw new InvalidOperationException("The workbook's Title built‑in property must not be empty before export.");

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
