using System;
using Aspose.Cells;

namespace AsposeCellsNumberFormatDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a style with a built‑in number format that shows two decimal places.
            // Style.Number = 2 corresponds to the "0.00" format.
            Style decimalStyle = workbook.CreateStyle();
            decimalStyle.Number = 2; // 0.00 format

            // Apply the style to the desired range (e.g., column A rows 1‑10)
            Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("A1:A10");
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            targetRange.ApplyStyle(decimalStyle, flag);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}