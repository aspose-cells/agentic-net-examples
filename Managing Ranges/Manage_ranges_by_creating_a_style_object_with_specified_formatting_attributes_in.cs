using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRangeStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data in the target range for visual reference
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["C2"].PutValue(0.5);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(20);
            cells["C3"].PutValue(0.3);

            // Create a style object and set desired formatting attributes
            Style style = workbook.CreateStyle();

            // Font settings
            style.Font.Name = "Calibri";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.White;

            // Cell background (solid fill)
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.DarkBlue;

            // Alignment
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.IsTextWrapped = true;

            // Create a range covering cells A1:C3
            var range = cells.CreateRange("A1", "C3");

            // Apply the style to the entire range
            range.SetStyle(style);

            // Save the workbook in XLSX format
            workbook.Save("StyledRange.xlsx");
        }
    }
}