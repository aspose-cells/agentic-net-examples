// Title: Apply the workbook's Dark2 theme color as the default border for every table in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: how to use Aspose.Cells C# to set a theme color as border for all ListObject cells | Aspose.Cells apply workbook Dark2 theme color to table borders in C# | set default thin border style for every table cell using Aspose.Cells .NET | retrieve Excel theme color with Aspose.Cells and apply to cell borders programmatically | iterate through worksheets and ListObjects to style borders with Aspose.Cells C#
// Tags: apply workbook theme color to ListObject borders Aspose.Cells | set thin border style for all table cells C# | retrieve Excel theme color using Aspose.Cells API | iterate worksheets and tables to modify cell borders Aspose | default border color for tables in Excel with Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // For ListObject
using Aspose.Cells.Drawing; // For BorderType and CellBorderType

// The C# program loads an existing Excel workbook, obtains the Dark2 theme color (fallback to Accent1), loops through each worksheet and its ListObjects, and assigns that theme color with a thin line style to the left, right, top, and bottom borders of every data cell before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook wb = new Workbook(inputPath);

            // Retrieve a theme color (using Accent1 as a fallback if Dark2 is unavailable)
            Color themeColor = wb.GetThemeColor(ThemeColorType.Accent1);

            // Iterate through all worksheets
            foreach (Worksheet ws in wb.Worksheets)
            {
                // Iterate through all tables (ListObjects) in the worksheet
                foreach (ListObject table in ws.ListObjects)
                {
                    // Get the data range of the table (excluding header)
                    Aspose.Cells.Range tblRange = table.DataRange;

                    // Loop through each cell in the table range
                    for (int row = tblRange.FirstRow; row < tblRange.FirstRow + tblRange.RowCount; row++)
                    {
                        for (int col = tblRange.FirstColumn; col < tblRange.FirstColumn + tblRange.ColumnCount; col++)
                        {
                            // Get the current cell style
                            Style cellStyle = ws.Cells[row, col].GetStyle();

                            // Apply the theme color to all four borders
                            cellStyle.Borders[BorderType.LeftBorder].Color = themeColor;
                            cellStyle.Borders[BorderType.RightBorder].Color = themeColor;
                            cellStyle.Borders[BorderType.TopBorder].Color = themeColor;
                            cellStyle.Borders[BorderType.BottomBorder].Color = themeColor;

                            // Set a thin line style for visibility
                            cellStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                            cellStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                            cellStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                            cellStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                            // Apply the modified style back to the cell
                            ws.Cells[row, col].SetStyle(cellStyle);
                        }
                    }
                }
            }

            // Save the modified workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
