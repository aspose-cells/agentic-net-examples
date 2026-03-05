using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDisplayRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some data to form a display range
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue(100);
            worksheet.Cells["B2"].PutValue(200);
            worksheet.Cells["A3"].PutValue(300);
            worksheet.Cells["B3"].PutValue(400);

            // Retrieve the maximum display range (includes data, merged cells, shapes)
            AsposeRange maxDisplayRange = worksheet.Cells.MaxDisplayRange;

            // If the worksheet is not empty, output range details
            if (maxDisplayRange != null)
            {
                Console.WriteLine("Max Display Range:");
                Console.WriteLine($"Start Row: {maxDisplayRange.FirstRow}");
                Console.WriteLine($"Start Column: {maxDisplayRange.FirstColumn}");
                Console.WriteLine($"Total Rows: {maxDisplayRange.RowCount}");
                Console.WriteLine($"Total Columns: {maxDisplayRange.ColumnCount}");

                // Create a new range that matches the max display range dimensions
                AsposeRange customRange = worksheet.Cells.CreateRange(
                    maxDisplayRange.FirstRow,
                    maxDisplayRange.FirstColumn,
                    maxDisplayRange.RowCount,
                    maxDisplayRange.ColumnCount);

                // Apply a simple style to the custom range
                Style style = workbook.CreateStyle();
                style.Font.IsBold = true;
                style.ForegroundColor = Color.LightGray;
                style.Pattern = BackgroundType.Solid;

                StyleFlag flag = new StyleFlag { FontBold = true, CellShading = true };
                customRange.ApplyStyle(style, flag);
            }

            // Save the workbook in XLSX format
            workbook.Save("DisplayRangeDemo.xlsx");
        }
    }
}