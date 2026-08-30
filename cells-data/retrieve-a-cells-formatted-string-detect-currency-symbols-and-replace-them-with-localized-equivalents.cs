// Title: Read a cell's formatted value and localize its currency symbol to the workbook's culture using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to obtain a cell's DisplayStringValue, detect any common currency symbol (e.g., $, €, £, ¥) and replace it with the CurrencySymbol of a provided CultureInfo, then write the localized string to another cell. | Create a reusable C# method that accepts a formatted numeric string and a CultureInfo, swaps known currency symbols with the culture's currency symbol, and demonstrate its integration in an Aspose.Cells workbook workflow.
// Common Searches: Aspose.Cells C# get formatted cell value and change currency symbol based on workbook culture | replace dollar sign with workbook CultureInfo currency symbol using Aspose.Cells | detect and localize currency symbols in Excel cell values with Aspose.Cells .NET | how to use DisplayStringValue to apply culture‑specific currency formatting in Aspose.Cells
// Tags: retrieve cell display string Aspose.Cells | currency symbol localization CultureInfo | replace known currency symbols formatted value C# | custom number format currency Aspose.Cells | write localized string to Excel cell Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCurrencyLocalization
{
    // Shows how to extract a cell's formatted string with Aspose.Cells, detect common currency symbols, replace them with the workbook's CultureInfo currency symbol, and write the localized result back to another cell in a .NET application.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a numeric value into cell A1
            Cell cell = cells["A1"];
            cell.PutValue(1234.56);

            // Apply a custom currency format that uses the dollar sign
            Style style = cell.GetStyle();
            style.Custom = "$#,##0.00";
            cell.SetStyle(style);

            // Retrieve the formatted string as displayed in Excel
            string formatted = cell.DisplayStringValue; // e.g., "$1,234.56"
            Console.WriteLine("Original formatted value: " + formatted);

            // Replace currency symbols with the symbols of the workbook's culture
            string localized = ReplaceCurrencySymbols(formatted, workbook.Settings.CultureInfo);
            Console.WriteLine("Localized formatted value: " + localized);

            // Optionally write the localized string back to another cell
            cells["B1"].PutValue(localized);
            cells["B1"].SetStyle(style); // keep the same number format for demonstration

            // Save the workbook
            workbook.Save("CurrencyLocalizationDemo.xlsx");
        }

        /// <param name="input">The formatted string containing a currency symbol.</param>
        /// <param name="culture">The target culture for localization.</param>
        /// <returns>The string with the currency symbol replaced.</returns>
        static string ReplaceCurrencySymbols(string input, CultureInfo culture)
        {
            // Get the currency symbol for the target culture
            string targetSymbol = culture.NumberFormat.CurrencySymbol;

            // Define a set of common currency symbols to replace
            string[] commonSymbols = new[] { "$", "€", "£", "¥", "₹", "₽", "₩", "₺", "₫", "₴", "₦", "₱", "₪", "₭", "₮", "₲", "₡", "₵", "₿" };

            // Replace any occurrence of a known symbol with the target symbol
            foreach (string symbol in commonSymbols)
            {
                if (input.Contains(symbol))
                {
                    input = input.Replace(symbol, targetSymbol);
                }
            }

            return input;
        }
    }
}
