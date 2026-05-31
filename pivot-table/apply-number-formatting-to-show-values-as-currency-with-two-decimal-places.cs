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

            // Put a numeric value into a cell
            Cell cell = sheet.Cells["B2"];
            cell.PutValue(1234.56);

            // Get the cell's style and set the built‑in currency format with two decimal places
            // According to Aspose.Cells documentation, Number = 7 corresponds to "$#,##0.00_);($#,##0.00)"
            Style style = cell.GetStyle();
            style.Number = 7;
            cell.SetStyle(style);

            // Save the workbook
            workbook.Save("CurrencyNumberFormat.xlsx");
        }
    }
}