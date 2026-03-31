using System;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source range A1:C3 with sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            // Apply a simple style to the source range
            Style style = workbook.CreateStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            AsposeRange sourceRange = sheet.Cells.CreateRange("A1:C3");
            sourceRange.SetStyle(style);

            // Define the destination range (starting at E5, same size as source)
            AsposeRange destRange = sheet.Cells.CreateRange(4, 4, 3, 3); // Row 5, Column 5 (zero‑based)

            // Copy source range to destination range preserving values and formatting
            sourceRange.Copy(destRange);

            // Save the workbook
            workbook.Save("RangeCopyPreserveFormatting.xlsx");
        }
    }
}