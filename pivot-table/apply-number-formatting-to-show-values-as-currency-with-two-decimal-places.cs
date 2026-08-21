// Title: Format Cells as Currency (Two Decimals) with Aspose.Cells for .NET
// Description: Creates a workbook, inserts numeric values into A1‑A3, applies the built‑in currency format (Number = 7, "$#,##0.00_);($#,##0.00)") to the range, and saves the file as CurrencyNumberFormatDemo.xlsx.
// Keywords: Aspose.Cells | .NET | C# | currency number format | built‑in format 7 | two decimal places | Excel export | apply style to range | format cells as money
// Common Searches: Aspose.Cells format cells as currency | C# apply built‑in number format 7 | how to set two‑decimal currency format in Aspose.Cells | apply number format to a range in .NET Excel | save workbook after currency formatting Aspose.Cells
// Developer Intent: Apply a built‑in currency format with two decimal places to a cell range and persist the workbook.
// Use Cases: Generate financial reports where all monetary columns display a consistent $‑style format. | Prepare data for pivot tables that require uniform currency representation. | Export application data to Excel while ensuring monetary values are correctly formatted.
// AI Prompts: Provide C# code that formats column B as currency with the local symbol and two decimal places using Aspose.Cells. | Show how to combine a custom currency format with conditional formatting for negative values in Aspose.Cells for .NET. | Generate a snippet that applies the built‑in currency style to multiple non‑contiguous ranges and saves the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsNumberFormattingDemo
{
    // Creates a workbook, inserts numeric values into A1‑A3, applies the built‑in currency format (Number = 7, "$#,##0.00_);($#,##0.00)") to the range, and saves the file as CurrencyNumberFormatDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put some numeric values in cells
                sheet.Cells["A1"].PutValue(1234.5);
                sheet.Cells["A2"].PutValue(5678.9);
                sheet.Cells["A3"].PutValue(9012.34);

                // Create a style and set the built‑in number format for currency with two decimals
                // Number = 7 corresponds to "$#,##0.00_);($#,##0.00)"
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Number = 7;

                // Apply the style to the range containing the values
                Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A3");
                range.ApplyStyle(currencyStyle, new StyleFlag { NumberFormat = true });

                // Save the workbook
                workbook.Save("CurrencyNumberFormatDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
