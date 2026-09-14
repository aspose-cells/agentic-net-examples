// Title: Throw a custom FreezeRowIndexOutOfRangeException when the requested freeze‑pane row exceeds the worksheet’s maximum row count in Aspose.Cells for .NET
// AI Prompts: Create a helper method that receives a Worksheet and a desired freeze‑row index, compares the index to Worksheet.Cells.MaxRow + 1, and throws a custom FreezeRowIndexOutOfRangeException with a clear message if the index is larger. | Integrate the validation method into existing code, then apply FreezePanes using the four‑parameter overload (row, column, rowSplit, columnSplit) only after the check passes, and save the workbook.
// Common Searches: Aspose.Cells how to check if freeze pane row index is within worksheet limits before calling FreezePanes | C# custom exception for out‑of‑range freeze row index in Aspose.Cells workbook | maximum row count returned by Worksheet.Cells.MaxRow for freeze pane validation Aspose.Cells | prevent ArgumentOutOfRangeException when setting freeze panes with large row numbers in Aspose.Cells .NET
// Tags: validate freeze pane row index Aspose.Cells | custom FreezeRowIndexOutOfRangeException C# | Worksheet.Cells.MaxRow row limit check | FreezePanes four‑parameter overload Aspose.Cells | handle invalid freeze row index .NET

using System;
using System.IO;
using Aspose.Cells;

// The example defines a custom FreezeRowIndexOutOfRangeException, validates that a requested freeze‑pane row does not exceed the worksheet's maximum row count (Cells.MaxRow + 1), throws the exception when the limit is breached, otherwise applies FreezePanes via the four‑parameter overload, and saves the workbook.
public class FreezeRowIndexOutOfRangeException : Exception
{
    public FreezeRowIndexOutOfRangeException(string message) : base(message) { }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook(); // placeholder for create rule
            Worksheet sheet = workbook.Worksheets[0];

            int requestedFreezeRow = 2_000_000; // Example index that may exceed limits

            // Apply freeze pane with validation
            ApplyFreezeRow(sheet, requestedFreezeRow);

            // Save the workbook (lifecycle rule)
            string outputPath = "output.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (FreezeRowIndexOutOfRangeException ex)
        {
            Console.WriteLine($"Freeze row error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    private static void ApplyFreezeRow(Worksheet sheet, int freezeRowIndex)
    {
        // Worksheet.Cells.MaxRow is zero‑based; maximum 1‑based row index = MaxRow + 1
        int maxAllowedRowIndex = sheet.Cells.MaxRow + 1;

        if (freezeRowIndex > maxAllowedRowIndex)
        {
            throw new FreezeRowIndexOutOfRangeException(
                $"Requested freeze row index {freezeRowIndex} exceeds worksheet maximum row count {maxAllowedRowIndex}.");
        }

        // Freeze rows above the specified index (column index 0 means no column freeze)
        // Use the 4‑parameter overload to avoid version‑specific issues
        sheet.FreezePanes(freezeRowIndex, 0, freezeRowIndex, 0);
    }
}
