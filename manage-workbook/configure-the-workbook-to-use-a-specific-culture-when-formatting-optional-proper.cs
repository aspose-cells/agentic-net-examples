using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook's culture to French (France)
            // This influences number, date, and custom format handling during export
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Optional: adjust decimal and group separators explicitly
            // (If not set, they are derived from the CultureInfo)
            workbook.Settings.NumberDecimalSeparator = ',';
            workbook.Settings.NumberGroupSeparator = ' ';

            // Add sample data
            Worksheet sheet = workbook.Worksheets[0];
            Cell cellNumber = sheet.Cells["A1"];
            cellNumber.PutValue(1234567.89); // Number to be formatted

            // Create a style with a custom numeric format
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00"; // Will use the culture's separators
            cellNumber.SetStyle(style);

            // Add a date value to demonstrate date formatting under the culture
            Cell cellDate = sheet.Cells["A2"];
            cellDate.PutValue(new DateTime(2023, 12, 31));
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd MMMM yyyy"; // French month names will be used
            cellDate.SetStyle(dateStyle);

            // Save the workbook (export) – format will respect the specified culture
            workbook.Save("CultureConfiguredWorkbook.xlsx");
        }
    }
}