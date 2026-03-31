using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill source range A1:C3 with sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            // Apply a simple style to the source range (bold header)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            // Create source range object (A1:C3)
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("A1:C3");
            // Apply style to the first row (header)
            sourceRange[0, 0].SetStyle(headerStyle);
            sourceRange[0, 1].SetStyle(headerStyle);
            sourceRange[0, 2].SetStyle(headerStyle);

            // Define destination range D5:F7 (same size as source)
            Aspose.Cells.Range destRange = sheet.Cells.CreateRange("D5:F7");

            // Copy data, formatting, and other cell properties from source to destination
            sourceRange.Copy(destRange);

            // Save the workbook to an XLSX file
            workbook.Save("RangeCopyWithFormatting.xlsx");
        }
    }
}