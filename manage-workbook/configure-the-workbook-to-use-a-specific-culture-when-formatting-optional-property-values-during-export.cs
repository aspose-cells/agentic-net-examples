using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Configure the workbook to use German (Germany) culture.
        // This affects number/date formatting when the workbook is exported.
        wb.Settings.CultureInfo = new CultureInfo("de-DE");

        // Create a style with a custom number format.
        // In German culture the decimal separator is a comma.
        Style style = wb.CreateStyle();
        style.Custom = "#,##0.00";

        // Apply the style to a cell and put a numeric value.
        Cell cell = wb.Worksheets[0].Cells["A1"];
        cell.PutValue(12345.67);
        cell.SetStyle(style);

        // Save the workbook; the formatted value respects the specified culture.
        wb.Save("CultureExportDemo.xlsx", SaveFormat.Xlsx);
    }
}