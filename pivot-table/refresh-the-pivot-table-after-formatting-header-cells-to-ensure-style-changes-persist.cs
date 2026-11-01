// Title: Refresh an Aspose.Cells PivotTable after applying a custom header style in C#
// Common Searches: C# Aspose.Cells how to keep pivot table header formatting after RefreshData | apply bold gray style to pivot table headers and persist with Aspose.Cells | refresh and recalculate pivot table after styling cells using Aspose.Cells .NET | Aspose.Cells pivot table header style not saved until table is refreshed | C# code to style pivot table header row and update pivot data
// Tags: Aspose.Cells pivot table header styling | refresh pivot table after style change | C# Aspose.Cells pivot table formatting | persist pivot table style changes | calculate pivot data after formatting

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot; // PivotTable class resides in this namespace

// The example loads an existing workbook, creates a bold light‑gray style, applies it to each header cell of the first pivot table, then calls RefreshData() and CalculateData() to ensure the formatting persists before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that a pivot table exists
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Determine the top‑left cell of the pivot table
            int headerRow = pivotTable.RowRange.StartRow;
            int headerColumn = pivotTable.RowRange.StartColumn;

            // Calculate the number of columns in the pivot table using its column range
            int columnCount = pivotTable.ColumnRange.EndColumn - pivotTable.ColumnRange.StartColumn + 1;

            // Create a style for the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to each header cell of the pivot table
            for (int col = 0; col < columnCount; col++)
            {
                try
                {
                    Cell cell = worksheet.Cells[headerRow, headerColumn + col];
                    cell.SetStyle(headerStyle);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to style cell at ({headerRow}, {headerColumn + col}): {ex.Message}");
                }
            }

            // Refresh and recalculate the pivot table to persist style changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
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
