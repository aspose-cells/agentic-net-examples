// Title: How to apply a light‑blue fill and black font to a pivot table header cell with Aspose.Cells for .NET
// Common Searches: c# asp.net set light blue background for pivot table header using Aspose.Cells | change font color of pivot table header cell programmatically with Aspose.Cells .NET | apply solid fill pattern to Excel pivot field header via Aspose.Cells library | style pivot table header cell in workbook with Aspose.Cells C# example | how to format pivot table header cell color in Aspose.Cells for .NET
// Tags: pivot header styling Aspose.Cells | cell fill color C# Aspose.Cells | solid background pattern Excel Aspose | format pivot field header .NET | light blue style Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.Drawing;
using System.IO;

// The code loads an existing workbook, locates the first pivot table's top‑left header cell, creates a style with a light‑blue solid fill and black font, applies this style to the header cell, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

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

            // Ensure a pivot table exists
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Determine header cell position (top‑left cell of the pivot table range)
            int headerRow = pivotTable.RowRange.StartRow;
            int headerColumn = pivotTable.ColumnRange.StartColumn;
            Cell headerCell = sheet.Cells[headerRow, headerColumn];

            // Create a style: light blue fill, black font
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            style.Font.Color = Color.Black;

            // Apply the style to the header cell
            headerCell.SetStyle(style);

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
