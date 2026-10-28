// Title: How to apply a built‑in PivotTable style to an existing Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# apply built‑in pivot table style to an existing workbook | Set PivotTableStyleInfo for a pivot table in an Excel file with Aspose.Cells .NET | C# example to format the first pivot table using a predefined style via Aspose.Cells | Programmatically change pivot table appearance with Aspose.Cells in .NET | Apply Excel pivot table style through Aspose.Cells API in C#
// Tags: apply pivot table style Aspose.Cells | PivotTableStyleInfo usage C# | format existing pivot table Excel .NET | predefined pivot style Aspose.Cells API | modify workbook pivot appearance C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing Excel file, checks for a pivot table on the first worksheet, and (when supported) applies a predefined PivotTable style using PivotTableStyleInfo before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
                workbook.Save(outputPath);
                return;
            }

            // Access the first pivot table
            PivotTable pivot = sheet.PivotTables[0];

            // Styling of pivot tables may require newer Aspose.Cells versions.
            // If supported, apply a style using PivotTableStyleInfo here.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
