// Title: How to apply a custom style to the totals row of an Aspose.Cells ListObject table in C#
// AI Prompts: Write C# code with Aspose.Cells that creates a worksheet, adds a ListObject table, enables the totals row, inserts SUM formulas, and applies a style featuring a light‑gray fill, bold dark‑blue font, and a thin bottom border to each totals cell. | Show how to build a reusable Style object in Aspose.Cells and assign it to every cell of a table's totals row after the row has been generated.
// Common Searches: aspnet c# apply custom formatting to totals row of an Excel table using Aspose.Cells | set background color and font style for totals row in Aspose.Cells ListObject | how to add bottom border to totals row cells in Aspose.Cells C# example | Aspose.Cells style totals row after enabling ShowTotals | C# code to style totals row of a table with light gray fill and dark blue bold text
// Tags: Aspose.Cells ListObject totals row styling | C# apply custom style to Excel table totals row | Aspose.Cells set cell background color and font | Excel table totals row formatting with Aspose.Cells | Create and style totals row in Aspose.Cells workbook

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds a ListObject table with sample product data, enables the totals row, inserts SUM formulas for Quantity and Price, defines a custom Style with a light‑gray background, bold dark‑blue font, and a thin black bottom border, applies this style to each cell in the totals row, and saves the file as TableWithStyledTotalsRow.xlsx.
    class ApplyTotalsRowStyle
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (A1:C5)
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["C3"].PutValue(0.3);

                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B4"].PutValue(15);
                sheet.Cells["C4"].PutValue(0.4);

                sheet.Cells["A5"].PutValue("Grape");
                sheet.Cells["B5"].PutValue(12);
                sheet.Cells["C5"].PutValue(0.6);

                // Define the range for the table (including header row)
                int firstRow = 0;   // zero‑based index for row 1
                int firstCol = 0;   // column A
                int totalRows = 5;  // header + 4 data rows
                int totalCols = 3;  // columns A‑C

                // Add a ListObject (table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(firstRow, firstCol,
                    firstRow + totalRows - 1, firstCol + totalCols - 1, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Enable the totals row
                table.ShowTotals = true;

                // Compute the index of the totals row (it is placed after the data rows)
                int totalsRowIndex = firstRow + totalRows; // zero‑based

                // Set the label for the first column in the totals row
                sheet.Cells[totalsRowIndex, firstCol].PutValue("Total");

                // Set SUM formulas for Quantity and Price columns
                // Quantity column (B)
                string qtyStart = sheet.Cells[firstRow + 1, firstCol + 1].Name; // B2
                string qtyEnd = sheet.Cells[firstRow + totalRows - 1, firstCol + 1].Name; // B5
                sheet.Cells[totalsRowIndex, firstCol + 1].Formula = $"=SUM({qtyStart}:{qtyEnd})";

                // Price column (C)
                string priceStart = sheet.Cells[firstRow + 1, firstCol + 2].Name; // C2
                string priceEnd = sheet.Cells[firstRow + totalRows - 1, firstCol + 2].Name; // C5
                sheet.Cells[totalsRowIndex, firstCol + 2].Formula = $"=SUM({priceStart}:{priceEnd})";

                // Create a custom style for the totals row
                Style totalsStyle = workbook.CreateStyle();
                totalsStyle.ForegroundColor = Color.LightGray;          // background color
                totalsStyle.Pattern = BackgroundType.Solid;
                totalsStyle.Font.IsBold = true;                         // bold font
                totalsStyle.Font.Color = Color.DarkBlue;                // font color
                totalsStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                totalsStyle.Borders[BorderType.BottomBorder].Color = Color.Black;

                // Apply the custom style to each cell in the totals row
                for (int col = firstCol; col < firstCol + totalCols; col++)
                {
                    Cell totalCell = sheet.Cells[totalsRowIndex, col];
                    totalCell.SetStyle(totalsStyle);
                }

                // Define output file path
                string outputPath = "TableWithStyledTotalsRow.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
