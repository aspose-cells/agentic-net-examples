// Title: Set Currency Number Format for Range R2:R20 with Aspose.Cells for .NET
// Description: Demonstrates how to create a style with the built‑in currency number format (Number = 5) in Aspose.Cells, apply it to the range R2:R20 on the first worksheet, and save the workbook as CurrencyRange.xlsx.
// Keywords: Aspose.Cells | C# | currency number format | Excel style | range R2:R20 | SetStyle | CreateStyle | built‑in format 5 | format cells as currency | Aspose.Cells tutorial
// Common Searches: Aspose.Cells set currency format C# | How to apply a number format to a range in Aspose.Cells | Apply built‑in currency style to cells R2 to R20 | C# Aspose.Cells format cells as currency | Create reusable style for multiple ranges Aspose.Cells
// Developer Intent: Apply a currency number format style to the cells R2‑R20.
// Use Cases: Format a column of monetary values (R2‑R20) as currency before exporting the workbook. | Maintain consistent financial formatting across several worksheets by reusing the same currency style. | Combine the currency style with additional formatting (fonts, borders, alignment) for polished financial reports.
// AI Prompts: Generate C# code that uses a custom currency format string instead of the built‑in Number = 5 in Aspose.Cells. | Show how to create a reusable currency style and apply it to multiple non‑contiguous ranges in a workbook using Aspose.Cells for .NET. | Explain how to merge a currency style with other style attributes such as font color, alignment, and borders in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a style with the built‑in currency number format (Number = 5) in Aspose.Cells, apply it to the range R2:R20 on the first worksheet, and save the workbook as CurrencyRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style and set its number format to a built‑in currency format (value 5)
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 5; // "$#,##0_);($#,##0)" – currency format

            // Define the range R2:R20 and apply the style
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("R2", "R20");
            range.SetStyle(currencyStyle);

            // Save the workbook
            workbook.Save("CurrencyRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
