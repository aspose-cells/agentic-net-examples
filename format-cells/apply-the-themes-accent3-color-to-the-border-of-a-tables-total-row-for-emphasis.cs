// Title: Apply the workbook’s Accent3 theme color to the borders of a table’s totals row using Aspose.Cells for .NET (C#)
// AI Prompts: Retrieve the workbook’s Accent3 theme color and assign it to the top, bottom, left, and right borders of each cell in the totals row of the first ListObject. | Enable the totals row of an Excel table and format all its cell borders with a thin line using the workbook’s Accent3 color via Aspose.Cells in C#. | Programmatically style the totals row of a ListObject by applying thin borders colored with the workbook’s Accent3 theme using Aspose.Cells for .NET.
// Common Searches: how to set theme accent color on Excel table totals row borders with Aspose.Cells C# | Aspose.Cells C# apply thin borders to ListObject totals row using workbook theme color | retrieve Accent3 color from workbook theme and style totals row in .NET | format totals row of an Excel table with theme colors using Aspose.Cells | C# code to highlight totals row borders with workbook Accent3 color in Aspose.Cells
// Tags: Aspose.Cells theme accent border styling for ListObject totals | C# thin border formatting on Excel table totals row | apply workbook Accent3 color to cell edges in totals row | Excel table totals row visual emphasis using Aspose.Cells | set border line style for ListObject totals row in .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, shows the totals row of the first table, obtains the workbook’s Accent3 theme color, and applies a thin border of that color to all four sides of each cell in the totals row before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one table on the sheet
                if (sheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables (ListObjects) found on the first worksheet.");
                    return;
                }

                // Assume the first table is the target
                ListObject table = sheet.ListObjects[0];

                // Show the totals row
                table.ShowTotals = true;

                // Determine the index of the totals row (first row after the data range)
                int totalRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;

                // Determine the first and last column indexes of the table's data range
                int firstColumn = table.DataRange.FirstColumn;
                int lastColumn = firstColumn + table.DataRange.ColumnCount - 1;

                // Retrieve the theme's Accent3 color
                Color accent3Color = workbook.GetThemeColor(ThemeColorType.Accent3);

                // Apply the Accent3 color to all borders of each cell in the totals row
                for (int col = firstColumn; col <= lastColumn; col++)
                {
                    Cell cell = sheet.Cells[totalRowIndex, col];
                    Style style = cell.GetStyle();

                    // Set border colors and line style for all four sides
                    style.Borders[BorderType.TopBorder].Color = accent3Color;
                    style.Borders[BorderType.BottomBorder].Color = accent3Color;
                    style.Borders[BorderType.LeftBorder].Color = accent3Color;
                    style.Borders[BorderType.RightBorder].Color = accent3Color;

                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    cell.SetStyle(style);
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
}
