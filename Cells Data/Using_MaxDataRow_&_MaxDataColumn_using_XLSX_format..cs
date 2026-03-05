using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsMaxDataDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(50);
            cells["C2"].PutValue(0.5);

            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(30);
            cells["C3"].PutValue(0.3);

            cells["A4"].PutValue("Cherry");
            cells["B4"].PutValue(20);
            cells["C4"].PutValue(1.2);

            // Retrieve the maximum data row and column indices (zero‑based)
            int maxDataRow = cells.MaxDataRow;       // Expected: 3 (row index of "Cherry")
            int maxDataColumn = cells.MaxDataColumn; // Expected: 2 (column index of "Price")

            Console.WriteLine($"MaxDataRow = {maxDataRow}");
            Console.WriteLine($"MaxDataColumn = {maxDataColumn}");

            // Create a range that covers all populated cells
            // Note: CreateRange expects the count of rows/columns, so add 1 to the max indices
            Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, maxDataRow + 1, maxDataColumn + 1);

            // Apply a simple style to the entire data range
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightYellow;
            style.Pattern = BackgroundType.Solid;

            StyleFlag flag = new StyleFlag
            {
                FontBold = true,
                CellShading = true
            };

            dataRange.ApplyStyle(style, flag);

            // Save the workbook as an XLSX file
            workbook.Save("MaxDataRowColumnDemo.xlsx");
        }
    }
}