using System;
using Aspose.Cells;

namespace AsposeCellsDateStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A (index 0) with sample date values
            cells["A1"].PutValue(new DateTime(2023, 1, 10));
            cells["A2"].PutValue(new DateTime(2023, 2, 15));
            cells["A3"].PutValue(new DateTime(2023, 3, 20));

            // Create a style and set a built‑in date number format (14 = "m/d/yyyy")
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // built‑in date format

            // Create a StyleFlag to apply only the number format
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the entire first column (column A)
            cells.ApplyColumnStyle(0, dateStyle, flag);

            // Save the workbook
            workbook.Save("DateStyleApplied.xlsx");
        }
    }
}