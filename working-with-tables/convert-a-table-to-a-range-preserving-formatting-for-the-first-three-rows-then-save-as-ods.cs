// Title: Convert a ListObject to a regular range while preserving the first three rows' formatting and save as ODS using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert ListObject to range preserving row styles | how to keep cell formatting when removing an Excel table with Aspose.Cells | save workbook as ODS after converting Excel table to range using Aspose.Cells | C# Aspose.Cells convert table to range without losing formatting | preserve first three rows formatting when converting ListObject to range Aspose.Cells
// Tags: ListObject to range conversion Aspose.Cells | retain row styles during table conversion | save workbook as ODS format C# | apply background style to rows Aspose.Cells | convert Excel table to regular range Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, fills a 5x3 range, adds a ListObject table, applies a light‑blue background to the first three rows, converts the table back to a normal range while keeping the applied styles, and finally saves the file as an ODS document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (5 rows x 3 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Add a table (ListObject) covering the data range A1:C5
            int firstRow = 0, firstCol = 0, totalRows = 5, totalCols = 3;
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol,
                firstRow + totalRows - 1, firstCol + totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table (Name property may not be available in some versions)
            table.DisplayName = "SampleTable";

            // Apply formatting to the first three rows of the table
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            StyleFlag flag = new StyleFlag { CellShading = true };

            for (int r = 0; r < 3; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    sheet.Cells[r, c].SetStyle(style, flag);
                }
            }

            // Convert the table back to a normal range while preserving formatting
            table.ConvertToRange();

            // Save the workbook as ODS
            workbook.Save("ConvertedTable.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
