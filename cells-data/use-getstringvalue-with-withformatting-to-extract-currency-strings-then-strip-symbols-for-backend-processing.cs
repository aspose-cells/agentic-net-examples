// Title: Retrieve a formatted currency string from an Excel cell with GetStringValue(DisplayString) and remove the monetary symbols for decimal conversion using Aspose.Cells for .NET
// AI Prompts: Demonstrate how to call GetStringValue with the DisplayString strategy to obtain the visible monetary text of a cell and then cleanse it of all non‑numeric characters in C#. | Generate C# code that captures the formatted amount from a worksheet cell, eliminates the currency sign, and converts the cleaned string to a decimal using Aspose.Cells.
// Common Searches: Aspose.Cells C# get visible currency text from cell without symbol | how to clean formatted Excel monetary value for backend processing | extract numeric portion from custom formatted cell using Aspose.Cells GetStringValue | convert Excel formatted amount to decimal in .NET with Aspose.Cells
// Tags: GetStringValue DisplayString monetary format | regex cleanse nonnumeric cell value | decimal conversion from formatted cell Aspose.Cells | custom monetary format extraction C# | server side numeric parsing Excel cell

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsCurrencyProcessing
{
    // Shows how to apply a custom currency format to a cell, retrieve the displayed string with GetStringValue(DisplayString), strip out all non‑numeric characters via regex, and parse the result into a decimal for backend use using Aspose.Cells for .NET.
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
            Style style = cell.GetStyle();
            style.Custom = "$#,##0.00";
            cell.SetStyle(style);

            // Extract the formatted string including the currency symbol
            // Using DisplayString to get exactly what Excel would show
            string formattedCurrency = cell.GetStringValue(CellValueFormatStrategy.DisplayString);
            Console.WriteLine("Formatted (with symbol): " + formattedCurrency);

            // Strip all non‑numeric characters (except decimal separator, comma and minus sign)
            string numericString = Regex.Replace(formattedCurrency, @"[^\d\.,-]+", "");
            Console.WriteLine("Stripped numeric string: " + numericString);

            // Optionally convert to a decimal for backend processing
            if (decimal.TryParse(numericString, out decimal numericValue))
            {
                Console.WriteLine("Parsed decimal value: " + numericValue);
            }
            else
            {
                Console.WriteLine("Failed to parse numeric value.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle)
            workbook.Save("CurrencyProcessingDemo.xlsx");
        }
    }
}
