using System;
using Aspose.Cells;

namespace AsposeCellsNumbersFormatDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Apply built‑in number formats using the Style.Number property
            ApplyNumberFormat(sheet, "A1", 1234.567, 0);   // General
            ApplyNumberFormat(sheet, "A2", 1234.567, 2);   // Decimal with two places
            ApplyNumberFormat(sheet, "A3", 1234.567, 5);   // Currency
            ApplyNumberFormat(sheet, "A4", 0.8567, 9);     // Percentage
            ApplyNumberFormat(sheet, "A5", DateTime.Now, 14); // Date (m/d/yyyy)
            ApplyNumberFormat(sheet, "A6", DateTime.Now, 21); // Time (h:mm:ss)

            // Apply a custom number format using the Style.Custom property
            ApplyCustomNumberFormat(sheet, "A7", 1234.567, "_-€ #,##0.00;[Red]-€ #,##0.00");

            // Save the workbook
            workbook.Save("NumbersFormatDemo.xlsx");
        }

        // Helper method to set a value and a built‑in number format
        static void ApplyNumberFormat(Worksheet sheet, string cellName, object value, int formatId)
        {
            Cell cell = sheet.Cells[cellName];
            cell.PutValue(value);

            // Retrieve the current style, set the built‑in format, and reapply
            Style style = cell.GetStyle();
            style.Number = formatId;
            cell.SetStyle(style);
        }

        // Helper method to set a value and a custom number format
        static void ApplyCustomNumberFormat(Worksheet sheet, string cellName, object value, string customFormat)
        {
            Cell cell = sheet.Cells[cellName];
            cell.PutValue(value);

            // Create a new style with the custom format and apply it
            Style style = sheet.Workbook.CreateStyle();
            style.Custom = customFormat;
            cell.SetStyle(style);
        }
    }
}