using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Set a specific CultureInfo for the workbook.
            // This culture will be used for date and number formatting.
            workbook.Settings.CultureInfo = new CultureInfo("de-DE"); // German uses comma as decimal separator

            // Access the first worksheet and a cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Put a numeric value that will be formatted according to the German culture
            cell.PutValue(1234567.89);

            // Create a style with a built‑in number format (e.g., Number with two decimal places)
            Style style = workbook.CreateStyle();
            style.Number = 2; // Built‑in format: "#,##0.00"

            // Apply the style to the cell
            cell.SetStyle(style);

            // Output the culture‑adjusted string representation of the cell value
            Console.WriteLine($"Formatted value (de-DE): {cell.StringValue}");

            // Save the workbook (uses the provided save rule)
            workbook.Save("CultureDemo.xlsx");
        }
    }
}