// Title: How to apply a custom light‑gray bold style to the totals row of a ListObject table using Aspose.Cells for C#
// AI Prompts: Write C# code that creates a ListObject, enables its totals row, adds SUM formulas for numeric columns, defines a Style with a solid light‑gray fill and bold black font, and applies that style to each cell in the totals row using Aspose.Cells. | Show step‑by‑step how to programmatically format the totals row of an Aspose.Cells table—setting background color, font weight, and applying the style after calculating column totals.
// Common Searches: Aspose.Cells C# format totals row of a ListObject with custom background color | C# apply bold font and gray fill to table totals row using Aspose.Cells | How to set style for totals row in an Aspose.Cells worksheet programmatically | Add SUM formulas and style the totals row in an Aspose.Cells table C#
// Tags: listobject totals row custom style Aspose.Cells | apply background color to totals row C# | bold font totals row Aspose.Cells | solid fill style for table rows Aspose.Cells | sum formula totals row Aspose.Cells | style totals row cells programmatically

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates creating a workbook, adding a ListObject table with a totals row, inserting SUM formulas for Quantity and Price columns, defining a light‑gray bold Style, and applying that Style to every cell in the totals row before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table (A1:C5)
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

            // Define the table range (A1:C5)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int dataRows = 5;          // number of data rows including header
            int totalColumns = 3;      // columns A‑C

            // Add a ListObject (table) and enable the totals row
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + dataRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";
            table.ShowTotals = true;

            // Configure totals row: label for first column, sum formulas for the others
            table.ListColumns[0].TotalsRowLabel = "Total";

            int totalsRowIndex = firstRow + dataRows; // zero‑based index of the totals row

            // Quantity column total (B)
            Cell qtyTotalCell = sheet.Cells[totalsRowIndex, firstColumn + 1];
            string qtyStart = sheet.Cells[firstRow + 1, firstColumn + 1].Name; // B2
            string qtyEnd = sheet.Cells[firstRow + dataRows - 1, firstColumn + 1].Name; // B5
            qtyTotalCell.Formula = $"=SUM({qtyStart}:{qtyEnd})";

            // Price column total (C)
            Cell priceTotalCell = sheet.Cells[totalsRowIndex, firstColumn + 2];
            string priceStart = sheet.Cells[firstRow + 1, firstColumn + 2].Name; // C2
            string priceEnd = sheet.Cells[firstRow + dataRows - 1, firstColumn + 2].Name; // C5
            priceTotalCell.Formula = $"=SUM({priceStart}:{priceEnd})";

            // Create a custom style for the totals row
            Style totalsStyle = workbook.CreateStyle();
            totalsStyle.ForegroundColor = Color.LightGray;
            totalsStyle.Pattern = BackgroundType.Solid;
            totalsStyle.Font.IsBold = true;
            totalsStyle.Font.Color = Color.Black;

            // Apply the style to each cell in the totals row
            for (int col = firstColumn; col < firstColumn + totalColumns; col++)
            {
                Cell cell = sheet.Cells[totalsRowIndex, col];
                cell.SetStyle(totalsStyle);
            }

            // Save the workbook
            workbook.Save("TableWithStyledTotals.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
