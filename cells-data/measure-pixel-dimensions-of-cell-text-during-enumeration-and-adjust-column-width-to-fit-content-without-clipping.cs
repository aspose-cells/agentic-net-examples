// Title: Auto‑fit Excel column widths by measuring cell text in pixels using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that iterates through each column of an Aspose.Cells worksheet, uses Cell.GetWidthOfValue to obtain the pixel width of each cell's displayed text, determines the maximum width per column, adds a small padding, and applies the result with Cells.SetColumnWidthPixel. | Show how to prevent text clipping in an Excel file by calculating pixel‑based column widths with Aspose.Cells and saving the workbook. | Provide a complete example that populates a worksheet with varied text lengths, measures pixel widths, auto‑adjusts column widths, and writes the file to disk.
// Common Searches: how to auto‑fit Excel column width by pixel in Aspose.Cells C# | measure displayed text width of a cell using Aspose.Cells | set column width in pixels with Aspose.Cells .NET example | avoid text clipping when exporting Excel with Aspose.Cells | C# calculate max pixel width per column Aspose.Cells
// Tags: Aspose.Cells GetWidthOfValue pixel measurement | Aspose.Cells SetColumnWidthPixel auto‑fit | C# compute max column pixel width | Excel column auto‑fit by pixel Aspose.Cells | prevent column text clipping C# Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook, fills cells with texts of different lengths, loops through each column to find the widest cell using Cell.GetWidthOfValue, adds a 5‑pixel padding, sets the column width with Cells.SetColumnWidthPixel, and saves the file as AutoFitByPixelDemo.xlsx.
public class AutoFitColumnByPixelDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in the worksheet
        cells["A1"].PutValue("Short");
        cells["A2"].PutValue("This is a considerably longer text that should expand the column");
        cells["B1"].PutValue("Another column");
        cells["B2"].PutValue("Medium length");
        cells["C1"].PutValue("A very very very long piece of text that will require a wide column");
        cells["C2"].PutValue("Tiny");

        // Determine the range of used rows and columns
        int maxRow = cells.MaxRow;       // zero‑based index of the last used row
        int maxColumn = cells.MaxColumn; // zero‑based index of the last used column

        // Iterate through each column to find the maximum pixel width of its cells
        for (int col = 0; col <= maxColumn; col++)
        {
            int maxPixelWidth = 0;

            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = cells[row, col];

                // Skip empty cells
                if (cell == null || cell.Type == CellValueType.IsNull)
                    continue;

                // Get the pixel width of the cell's displayed value
                int cellPixelWidth = cell.GetWidthOfValue();

                // Track the widest cell in the column
                if (cellPixelWidth > maxPixelWidth)
                    maxPixelWidth = cellPixelWidth;
            }

            // Add a small padding to avoid clipping (e.g., 5 pixels)
            int paddedWidth = maxPixelWidth + 5;

            // Set the column width in pixels
            cells.SetColumnWidthPixel(col, paddedWidth);
        }

        // Save the workbook with the adjusted column widths
        workbook.Save("AutoFitByPixelDemo.xlsx");
        Console.WriteLine("Workbook saved as AutoFitByPixelDemo.xlsx");
    }
}
