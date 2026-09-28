// Title: How to reset worksheet panes and freeze the top row using Aspose.Cells for .NET (C#)
// AI Prompts: Provide C# code that invokes Worksheet.ResetPanes to clear any existing splits, then freezes the first row with Worksheet.FreezePanes using Aspose.Cells. | Demonstrate how to load a workbook, reset its panes, and set a new top‑row freeze pane in a .NET application with Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to clear worksheet splits before freezing rows | reset worksheet panes then freeze first row using C# Aspose.Cells | example of Worksheet.ResetPanes followed by FreezePanes in Excel file
// Tags: Worksheet.ResetPanes Aspose.Cells | reset worksheet panes .NET | freeze top row Worksheet.FreezePanes | clear Excel splits C# Aspose | apply new freeze pane after reset

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, calls Worksheet.ResetPanes to remove all existing splits and freeze panes, then uses Worksheet.FreezePanes to lock the first row, and finally saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Clear any existing freeze panes by resetting them to default (0,0)
            sheet.FreezePanes(0, 0, 0, 0);

            // Apply a new pane configuration: freeze the first row
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
