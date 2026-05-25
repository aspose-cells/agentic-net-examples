using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsCurrencyProcessing
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put a numeric value that represents an amount
                Cell amountCell = cells["B2"];
                amountCell.PutValue(1234.56);

                // Apply a built‑in currency format (Number format ID 44 = "$#,##0.00")
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Number = 44; // Currency format
                amountCell.SetStyle(currencyStyle);

                // Extract the formatted string using GetStringValue with DisplayString strategy
                // This returns exactly what Excel would display (e.g., "$1,234.56")
                string formattedCurrency = amountCell.GetStringValue(CellValueFormatStrategy.DisplayString);
                Console.WriteLine($"Formatted currency string: {formattedCurrency}");

                // Strip currency symbols and grouping separators, keep digits, decimal separator and sign
                NumberFormatInfo nfi = CultureInfo.CurrentCulture.NumberFormat;
                string decimalSeparator = Regex.Escape(nfi.NumberDecimalSeparator);
                // Build a regex that removes everything except digits, sign, and decimal separator
                string cleaned = Regex.Replace(formattedCurrency, $"[^0-9\\-+{decimalSeparator}]", string.Empty);
                Console.WriteLine($"Cleaned numeric string: {cleaned}");

                // Optionally convert the cleaned string back to a numeric type for backend processing
                if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.CurrentCulture, out double numericValue))
                {
                    Console.WriteLine($"Numeric value for backend: {numericValue}");
                }
                else
                {
                    Console.WriteLine("Failed to parse the cleaned string to a numeric value.");
                }

                // Save the workbook (optional, just to demonstrate lifecycle compliance)
                string outputPath = "CurrencyProcessingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}