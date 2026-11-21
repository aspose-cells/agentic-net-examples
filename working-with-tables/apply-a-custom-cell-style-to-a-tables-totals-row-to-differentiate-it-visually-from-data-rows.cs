// Title: How to apply a custom style to the totals row of an Excel table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# apply background color to table totals row | Set bold font for Excel ListObject totals row with Aspose.Cells | Change style of totals row in an Excel table programmatically using Aspose.Cells .NET | Customize totals row appearance in a workbook loaded with Aspose.Cells | Apply solid fill style to totals row cells in an Excel table via Aspose.Cells C#
// Tags: custom ListObject totals row style Aspose.Cells | solid fill totals row cells C# | bold font totals row formatting | Excel table totals row appearance Aspose.Cells | C# style totals row in workbook

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing workbook, checks for a table with a visible totals row, creates a style with a light gray solid background and bold black font, applies this style to each cell in the totals row, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            ListObject table = sheet.ListObjects[0];

            // Verify that the table displays a totals row
            if (!table.ShowTotals)
            {
                Console.WriteLine("The table does not have a totals row.");
                return;
            }

            // Get the index of the totals row (row after the data range)
            int totalsRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;

            // Determine the column range of the table's data (excluding header/totals)
            int firstCol = table.DataRange.FirstColumn;
            int lastCol = firstCol + table.DataRange.ColumnCount - 1;

            // Create a custom style for the totals row
            Style totalsStyle = workbook.CreateStyle();
            totalsStyle.ForegroundColor = Color.LightGray;
            totalsStyle.Pattern = BackgroundType.Solid;
            totalsStyle.Font.IsBold = true;
            totalsStyle.Font.Color = Color.Black;

            // Apply the style to each cell in the totals row within the table range
            for (int col = firstCol; col <= lastCol; col++)
            {
                Cell cell = sheet.Cells[totalsRowIndex, col];
                cell.SetStyle(totalsStyle);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
