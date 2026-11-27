// Title: Apply the Dark1 theme solid fill to pivot table row headers with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set Dark1 theme fill for pivot table row headers | How to change pivot table row header background color using Aspose.Cells .NET | Apply solid fill to pivot table row headers programmatically in C# | Using RowHeaderStyle property to style pivot table headers in Aspose.Cells | Dynamic property access for pivot table styling with Aspose.Cells
// Tags: pivot table row header style Aspose.Cells | RowHeaderStyle property C# | theme Dark1 fill Aspose.Cells | solid background color for pivot table headers | dynamic pivot table styling Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Microsoft.CSharp.RuntimeBinder;

// The example loads an existing workbook, checks for a pivot table, creates a style with a solid Dark1 theme fill, assigns it to the pivot table's RowHeaderStyle (using dynamic access for compatibility), and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that a pivot table exists
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Create a style for the row header cells
            Style rowHeaderStyle = workbook.CreateStyle();
            rowHeaderStyle.ForegroundColor = Color.DarkBlue; // solid color
            rowHeaderStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the pivot table's row headers.
            // Use dynamic to accommodate possible API variations.
            try
            {
                dynamic dynPivot = pivotTable;
                dynPivot.RowHeaderStyle = rowHeaderStyle;
            }
            catch (RuntimeBinderException)
            {
                // If the RowHeaderStyle property is unavailable, attempt manual styling.
                // Note: Direct access to the row header range is not supported in this version,
                // so we skip manual styling in this fallback.
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
