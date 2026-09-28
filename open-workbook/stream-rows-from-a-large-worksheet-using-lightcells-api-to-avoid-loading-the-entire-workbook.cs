// Title: Load a large XLSX workbook in C# using Aspose.Cells LoadOptions (basic example)
// AI Prompts: Write C# code that checks if an Excel file exists and then opens it with Aspose.Cells using a specific LoadOptions instance. | Show how to catch and log exceptions when initializing a Workbook object for a large XLSX file in .NET. | Demonstrate creating a LoadOptions object for the XLSX format and passing it to the Workbook constructor in C#.
// Common Searches: Aspose.Cells C# load large XLSX file with LoadOptions example | how to verify Excel file path before opening workbook in Aspose.Cells .NET | exception handling for Workbook constructor Aspose.Cells C#
// Tags: Aspose.Cells LoadOptions for XLSX format | C# verify Excel file existence before loading | exception handling when opening workbook Aspose.Cells | large workbook loading with Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that 'LargeWorkbook.xlsx' exists, creates a LoadOptions object for the XLSX format, loads the workbook with Aspose.Cells, and prints a success message or any caught exception.
class Program
{
    static void Main()
    {
        // Path to the large Excel file
        string filePath = "LargeWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File '{filePath}' not found.");
            return;
        }

        try
        {
            // Load the workbook using default LoadOptions (no custom LightCells provider)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Workbook loaded successfully; further processing can be added here if needed
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
