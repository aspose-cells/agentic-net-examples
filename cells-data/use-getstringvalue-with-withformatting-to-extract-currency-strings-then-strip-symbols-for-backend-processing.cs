// Title: Extract a formatted currency string and convert to numeric value with Aspose.Cells for .NET
// Description: Creates a workbook, writes a numeric amount, applies a custom currency format, retrieves the display string via GetStringValue(CellValueFormatStrategy.DisplayString), strips non‑numeric characters with Regex, and parses the result to a decimal for backend processing before saving the file.
// Keywords: Aspose.Cells GetStringValue | currency format extraction .NET | CellValueFormatStrategy DisplayString | remove currency symbol C# | regex numeric string conversion | parse decimal from cell value
// Common Searches: Aspose.Cells get formatted currency string | strip currency symbol from cell value C# | convert cell display string to decimal Aspose | GetStringValue DisplayString example | regex clean numeric string from formatted value
// Developer Intent: Retrieve a cell's formatted currency text, remove symbols, and convert it to a numeric type for further processing.
// Use Cases: Call GetStringValue with CellValueFormatStrategy.DisplayString to obtain the cell's visible currency text. | Use a regular expression to eliminate all characters except digits, decimal separators, and sign symbols. | Parse the cleaned string to a decimal with decimal.TryParse for calculations or storage.
// AI Prompts: Show how to use Aspose.Cells GetStringValue(DisplayString) to get a currency formatted string and convert it to a decimal in C#. | Provide a C# snippet that extracts a formatted monetary value from a worksheet cell, removes currency symbols and grouping separators using Regex, and safely parses it to a decimal.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsCurrencyExample
{
    // Creates a workbook, writes a numeric amount, applies a custom currency format, retrieves the display string via GetStringValue(CellValueFormatStrategy.DisplayString), strips non‑numeric characters with Regex, and parses the result to a decimal for backend processing before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put a numeric value that represents an amount
            Cell amountCell = cells["A1"];
            amountCell.PutValue(1234.56);

            // Apply a custom currency format (e.g., $1,234.56)
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";
            amountCell.SetStyle(currencyStyle);

            // Extract the formatted string including the currency symbol
            string formattedWithSymbol = amountCell.GetStringValue(CellValueFormatStrategy.DisplayString);
            Console.WriteLine("Formatted with symbol: " + formattedWithSymbol);

            // Strip all non‑numeric characters except decimal separator and minus sign
            string numericString = Regex.Replace(formattedWithSymbol, @"[^\d\.,-]+", "");
            Console.WriteLine("Numeric string for backend: " + numericString);

            // Optionally convert to a decimal for further processing
            if (decimal.TryParse(numericString, out decimal amount))
            {
                Console.WriteLine("Parsed decimal value: " + amount);
            }
            else
            {
                Console.WriteLine("Failed to parse numeric string.");
            }

            // Save the workbook (lifecycle rule: create → save)
            workbook.Save("CurrencyExample.xlsx");
        }
    }
}
