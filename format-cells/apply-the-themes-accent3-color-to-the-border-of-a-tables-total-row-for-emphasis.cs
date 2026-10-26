// Title: How to apply the workbook’s Accent3 theme color to the borders of a table’s total row using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set Accent3 theme color on table total row borders | how to change the border color of a ListObject total row using Aspose.Cells | retrieve workbook theme accent color and apply to Excel table total row in .NET | C# example for formatting total row borders with workbook theme in Aspose.Cells | apply theme color to total row border of an Excel table with Aspose.Cells for .NET
// Tags: theme accent border styling Aspose.Cells | ListObject total row border style C# | retrieve workbook theme color Aspose.Cells | set cell border style for table total row using Aspose.Cells | Aspose.Cells border color customization

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook, obtains the Accent3 color from the workbook’s theme, accesses the first ListObject, ensures its total row is visible, creates a range for that total row, builds a style that sets all four borders to the Accent3 color, applies the style to each cell in the total row, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook wb = new Workbook(inputPath);

            // Retrieve the theme's Accent3 color
            Color accent3 = wb.GetThemeColor(ThemeColorType.Accent3);

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Ensure the worksheet contains at least one table
            if (ws.ListObjects.Count > 0)
            {
                // Get the first table in the worksheet
                ListObject table = ws.ListObjects[0];

                // Make sure the total row is visible
                table.ShowTotals = true;

                // Determine the range representing the total row
                int totalRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount; // row after data rows
                int firstCol = table.DataRange.FirstColumn;
                int colCount = table.DataRange.ColumnCount;
                AsposeRange totalRow = ws.Cells.CreateRange(totalRowIndex, firstCol, 1, colCount);

                // Create a style that applies Accent3 to all borders
                Style borderStyle = wb.CreateStyle();
                borderStyle.Borders[BorderType.TopBorder].Color = accent3;
                borderStyle.Borders[BorderType.BottomBorder].Color = accent3;
                borderStyle.Borders[BorderType.LeftBorder].Color = accent3;
                borderStyle.Borders[BorderType.RightBorder].Color = accent3;

                // Apply the style to each cell in the total row
                foreach (Cell cell in totalRow)
                {
                    StyleFlag flag = new StyleFlag { All = true };
                    cell.SetStyle(borderStyle, flag);
                }
            }

            // Save the modified workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
