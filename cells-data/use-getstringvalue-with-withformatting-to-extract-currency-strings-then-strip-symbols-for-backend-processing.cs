using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsCurrencyExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a numeric value into cell A1
            Cell cell = cells["A1"];
            cell.PutValue(1234.56);

            // Apply a custom currency format (e.g., $1,234.56)
            Style style = workbook.CreateStyle();
            style.Custom = "$#,##0.00";
            cell.SetStyle(style);

            // Get the formatted string value as it appears in Excel (includes the currency symbol)
            string formattedCurrency = cell.GetStringValue(CellValueFormatStrategy.DisplayString);
            Console.WriteLine("Formatted (with currency symbol): " + formattedCurrency);

            // Strip all non-numeric characters except decimal separator and minus sign
            // This yields a plain numeric string suitable for backend processing
            string numericString = Regex.Replace(formattedCurrency, @"[^\d\.\-]", "");
            Console.WriteLine("Numeric string (symbols removed): " + numericString);

            // Optionally, convert the numeric string to a decimal for further calculations
            if (decimal.TryParse(numericString, out decimal numericValue))
            {
                Console.WriteLine("Parsed decimal value: " + numericValue);
            }

            // Save the workbook to a file
            workbook.Save("CurrencyExample.xlsx");
        }
    }
}