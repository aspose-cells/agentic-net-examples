// Title: Delete rows above a freeze pane and reset FreezePanes in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Remove all rows preceding a specified freeze row, then reapply Worksheet.FreezePanes so the header stays frozen in a C# Aspose.Cells workbook. | After deleting rows above the original freeze point, compute the new freeze row index and invoke Worksheet.FreezePanes with those values in Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells delete rows above freeze pane and keep header frozen | How to adjust FreezePanes after removing rows in an Excel workbook using Aspose.Cells | Aspose.Cells .NET programmatically delete top rows and reset freeze panes
// Tags: remove top rows before freeze pane Aspose.Cells C# | reset FreezePanes after row deletion Aspose.Cells | Worksheet.FreezePanes overload usage .NET | adjust freeze row index after rows removed Excel

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook, deletes all rows above a specified freeze row, recalculates the freeze row index, applies Worksheet.FreezePanes with the new parameters, and saves the modified file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Original freeze point (zero‑based row index). For example, freeze at row 6 (1‑based).
            int originalFreezeRow = 5;

            // Delete all rows above the original freeze point.
            if (originalFreezeRow > 0)
            {
                sheet.Cells.DeleteRows(0, originalFreezeRow);
            }

            // After deletion the visual freeze point shifts upward.
            // Keep the first row visible (header) after deletion.
            int newFreezeRow = 1; // freeze rows above row index 1

            // Apply the new FreezePanes settings.
            // Use Worksheet.FreezePanes overload (compatible with all Aspose.Cells versions).
            sheet.FreezePanes(newFreezeRow, 0, newFreezeRow, 0);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
