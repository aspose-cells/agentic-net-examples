using System;
using Aspose.Cells;

namespace AsposeCellsNumberFormatDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234.56);

            // Create a style and set built‑in currency format with two decimal places
            // According to Aspose.Cells documentation, Number = 7 corresponds to "$#,##0.00_);($#,##0.00)"
            Style style = workbook.CreateStyle();
            style.Number = 7;

            // Apply the style to the cell
            cell.SetStyle(style);

            // Save the workbook
            workbook.Save("CurrencyNumberFormat.xlsx");
        }
    }
}