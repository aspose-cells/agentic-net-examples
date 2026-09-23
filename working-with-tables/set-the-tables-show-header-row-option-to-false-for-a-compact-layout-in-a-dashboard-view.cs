// Title: Hide the header row and apply a compact dashboard style to an Excel ListObject with Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing workbook, access the first ListObject, set ShowHeaderRow to false, turn off ShowTableStyleFirstColumn, ShowTableStyleLastColumn, ShowTableStyleRowStripes, and ShowTableStyleColumnStripes, then save the workbook. | Programmatically remove table decorations (header row, first/last column highlights, row and column stripes) to create a minimal dashboard view using Aspose.Cells in C#.
// Common Searches: asp.net hide ListObject header row using Aspose.Cells | create compact dashboard table layout in Excel with Aspose.Cells C# | disable first and last column style for Excel table Aspose.Cells | remove row and column stripe formatting from Excel table Aspose.Cells | how to turn off table style elements in Aspose.Cells .NET
// Tags: hide ListObject header Aspose.Cells | compact dashboard table styling Aspose.Cells | disable first column table style Aspose.Cells | remove row stripe formatting Aspose.Cells | modify Excel table properties C# | Aspose.Cells table style customization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads an existing Excel workbook, accesses the first worksheet's first ListObject, disables the header row and all table style decorations (first/last column indicators and row/column stripes) to produce a compact dashboard appearance, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Access the first table
            ListObject table = sheet.ListObjects[0];

            // Hide the header row
            table.ShowHeaderRow = false;

            // Compact layout: hide first/last column style indicators
            table.ShowTableStyleFirstColumn = false;
            table.ShowTableStyleLastColumn = false;

            // Dashboard view: hide row/column stripes
            table.ShowTableStyleRowStripes = false;
            table.ShowTableStyleColumnStripes = false;

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
