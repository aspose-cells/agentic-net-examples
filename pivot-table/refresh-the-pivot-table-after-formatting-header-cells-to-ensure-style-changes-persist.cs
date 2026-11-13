// Title: How to refresh an Aspose.Cells PivotTable after applying a custom header style in C#
// Common Searches: Aspose.Cells C# refresh pivot table after setting header cell style | preserve custom header formatting in Aspose.Cells pivot tables | apply style to pivot table header row and refresh data using Aspose.Cells | C# code to style pivot table headers and call RefreshData with Aspose.Cells | Aspose.Cells pivot table header style not retained without RefreshData
// Tags: Aspose.Cells RefreshData pivot table | C# Aspose.Cells header style pivot table | pivot table header formatting Aspose.Cells | Aspose.Cells workbook save after pivot refresh | apply custom style to pivot table header C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The example loads an existing workbook, checks for a pivot table on the first worksheet, creates a custom style (white bold font on dark blue background), applies the style to the pivot table's header row cells, calls RefreshData to update cached data and keep the formatting, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No pivot tables were found on the first worksheet.");

            // Retrieve the first pivot table on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Create a style for the header cells
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.White;
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;

            // Determine the header row range of the pivot table (first row of the table range)
            CellArea tableRange = pivotTable.TableRange1;
            int headerRow = tableRange.StartRow;
            int startCol = tableRange.StartColumn;
            int endCol = tableRange.EndColumn;

            // Apply the style to each cell in the header row
            for (int col = startCol; col <= endCol; col++)
            {
                Cell cell = sheet.Cells[headerRow, col];
                cell.SetStyle(headerStyle);
            }

            // Refresh the pivot table to ensure any cached data is up‑to‑date
            pivotTable.RefreshData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
