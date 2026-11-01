// Title: How to clear all page fields from a PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove all page filters from a PivotTable | How to delete page fields in an Excel PivotTable using Aspose.Cells | Programmatically clear PivotTable page items with Aspose.Cells for .NET | Example code to reset PivotTable page fields in a workbook using Aspose | Clear PivotTable page fields before saving workbook with Aspose.Cells
// Tags: Aspose.Cells PivotTable PageFields.Clear method | C# remove pivot page filters with Aspose | Excel workbook update pivot page fields Aspose.Cells | Programmatic reset of PivotTable page items Aspose | Aspose.Cells purge pivot page fields

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for a PivotTable on the first worksheet, clears all page fields from that PivotTable using the PageFields.Clear method, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Get the first PivotTable in the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Clear all page fields from the PivotTable
            pivotTable.PageFields.Clear();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
