using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizedNumberFormat
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with mixed content
            cells["A1"].PutValue(1234.56);          // numeric
            cells["A2"].PutValue("Sample Text");   // text
            cells["A3"].PutValue(98765);           // integer
            cells["A4"].PutValue(0.789);           // numeric
            cells["B1"].PutValue(DateTime.Now);    // date (treated as numeric for formatting)

            // Set an initial generic number format for demonstration
            Style initStyle = workbook.CreateStyle();
            initStyle.Custom = "#,##0.00";
            cells["A1"].SetStyle(initStyle);
            cells["A3"].SetStyle(initStyle);
            cells["A4"].SetStyle(initStyle);
            cells["B1"].SetStyle(initStyle);

            // Change workbook culture to French (France) to obtain localized number formats
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Iterate through used cells and replace numeric patterns with localized formats
            foreach (Cell cell in cells)
            {
                // Process only cells that contain numeric values (including dates)
                if (cell.IsNumericValue)
                {
                    // Retrieve the current display string (formatted according to current style)
                    string currentDisplay = cell.DisplayStringValue;

                    // Parse the underlying numeric value using invariant culture
                    // This works for both double and DateTime (DateTime is stored as double internally)
                    double numericValue = cell.DoubleValue;

                    // Format the numeric value using the workbook's current culture
                    // The "N" format specifier respects the culture's number group and decimal separators
                    string localizedString = numericValue.ToString("N", workbook.Settings.CultureInfo);

                    // Apply a custom culture-dependent format to the cell style
                    // Setting CultureCustom ensures the pattern adapts to the workbook's culture
                    Style style = cell.GetStyle();
                    style.CultureCustom = "#,##0.00"; // pattern stays the same, separators become culture-specific
                    cell.SetStyle(style);

                    // Optionally, you can verify the replacement by writing to console
                    Console.WriteLine($"Cell {cell.Name}: Original='{currentDisplay}' => Localized='{cell.DisplayStringValue}'");
                }
                else
                {
                    // Non-numeric cells remain unchanged; display their original string
                    Console.WriteLine($"Cell {cell.Name}: Non-numeric value = '{cell.DisplayStringValue}'");
                }
            }

            // Save the modified workbook (lifecycle: save)
            workbook.Save("LocalizedNumbers.xlsx");
        }
    }
}