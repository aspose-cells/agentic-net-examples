using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Configure the workbook to use a specific culture (German - Germany)
        // This affects date and number formatting throughout the workbook
        workbook.Settings.CultureInfo = new CultureInfo("de-DE");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Demonstrate number formatting with the specified culture
        Cell numberCell = sheet.Cells["A1"];
        numberCell.PutValue(1234567.89); // Put a numeric value

        // Apply a built‑in number format (currency) which will be rendered using German conventions
        Style numberStyle = numberCell.GetStyle();
        numberStyle.Number = 4; // Currency format
        numberCell.SetStyle(numberStyle);

        // Demonstrate date formatting with the specified culture
        Cell dateCell = sheet.Cells["A2"];
        dateCell.PutValue(DateTime.Now); // Put the current date/time

        // Apply a built‑in date format which will be rendered using German conventions
        Style dateStyle = dateCell.GetStyle();
        dateStyle.Number = 14; // Short date format
        dateCell.SetStyle(dateStyle);

        // Save the workbook to a file
        workbook.Save("LocalizedWorkbook.xlsx");
    }
}