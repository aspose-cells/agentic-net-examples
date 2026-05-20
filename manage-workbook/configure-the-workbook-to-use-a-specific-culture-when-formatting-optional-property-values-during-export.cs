using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Set the desired culture for formatting optional property values
            // For example, French (France) uses space as group separator and comma as decimal separator
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Optional: demonstrate the effect by adding a number and applying a custom style
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234567.89); // Value to be formatted

            // Create a style with a custom numeric format
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00"; // Uses the workbook's culture settings
            cell.SetStyle(style);

            // Save the workbook (save rule)
            workbook.Save("Workbook_With_FrenchCulture.xlsx");
        }
    }
}