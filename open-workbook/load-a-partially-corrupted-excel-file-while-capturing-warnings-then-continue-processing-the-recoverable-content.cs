// Title: Load a partially corrupted Excel workbook in C# with Aspose.Cells while capturing load warnings and processing recoverable data
// Common Searches: C# Aspose.Cells load corrupted Excel file and retrieve load warnings | How to open a partially damaged .xlsx using Aspose.Cells LoadOptions | Recover data from a corrupted workbook with Aspose.Cells in .NET | Capture and log load warnings when loading an Excel file with Aspose.Cells | Continue processing after loading a corrupted workbook using Aspose.Cells
// Tags: Aspose.Cells load corrupted xlsx | recoverable content Aspose.Cells workbook | load options warnings Aspose.Cells | handle workbook load exceptions C# | open partially damaged Excel file .NET

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the Excel file exists, creates a default LoadOptions object, and loads the workbook with Aspose.Cells, which automatically recovers readable content from a partially corrupted file. It then accesses the first worksheet, prints its name and the value of cell A1, and wraps the operation in a try‑catch block to capture any load warnings or exceptions.
class Program
{
    static void Main()
    {
        // Path to the potentially corrupted workbook.
        string filePath = "CorruptedFile.xlsx";

        // Verify that the file exists to avoid FileNotFoundException.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        try
        {
            // Configure load options (auto-detect format, default behavior loads recoverable content).
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the first worksheet and read a cell value.
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"First sheet name: {sheet.Name}");
            Console.WriteLine($"Value in A1: {sheet.Cells["A1"].StringValue}");

            // Continue with further processing as needed...
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions (e.g., loading errors).
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
