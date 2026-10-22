// Title: Get and log the unique worksheet index (SheetId) from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that checks for an .xlsx file, loads it with Aspose.Cells, reads the Index property of the first worksheet, and writes the worksheet name and index to the console for audit purposes. | Show how to use Aspose.Cells Workbook and Worksheet objects to retrieve a worksheet's SheetId and output it as a log entry in a .NET application.
// Common Searches: asp.net retrieve worksheet index with Aspose.Cells for audit logging | c# how to get SheetId of an Excel sheet using Aspose.Cells | example code to log worksheet unique identifier from a workbook in .NET | Aspose.Cells get worksheet Index property after loading .xlsx file | verify file existence before reading worksheet SheetId with Aspose.Cells
// Tags: Aspose.Cells worksheet identifier extraction | C# log Excel sheet ID | audit trail worksheet SheetId | load workbook read sheet Index .NET | pre‑load file existence validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that 'input.xlsx' exists, loads it with Aspose.Cells, accesses the first worksheet, reads its Index property (the unique SheetId within the workbook), and writes the worksheet name and index to the console while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the worksheet's index (unique within the workbook)
            int sheetIndex = sheet.Index;

            // Log the index for audit tracking
            Console.WriteLine($"Worksheet '{sheet.Name}' has Index: {sheetIndex}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
