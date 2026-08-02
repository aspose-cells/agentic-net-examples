// Title: Aspose.Cells for .NET – Apply a Custom Currency Format ("$#,##0.00") with Two Decimals to a Cell Range (C#)
// Description: Creates a new workbook, fills cells A1‑A5 with sample numbers, defines a style using the custom format "$#,##0.00", applies only the number‑format part via StyleFlag to the range, and saves the file as CurrencyNumberFormat.xlsx.
// Keywords: Aspose.Cells C# | custom currency number format | apply number format to range | StyleFlag number format | currency formatting Aspose.Cells | two decimal places Excel | CurrencyNumberFormat.xlsx | financial spreadsheet styling | US dollar format Aspose
// Common Searches: Aspose.Cells set custom currency format for a range | How to use StyleFlag to change only number format in Aspose.Cells | C# code to format cells as $#,##0.00 with Aspose.Cells | Apply two‑decimal currency style to multiple cells Aspose | Save workbook after applying currency format Aspose.Cells
// Developer Intent: Apply a two‑decimal currency style to a specific range of cells without affecting other cell properties.
// Use Cases: Standardize monetary columns in financial reports. | Generate invoices where every amount shows a dollar sign and two decimal places. | Present sales data with consistent currency formatting across a table.
// AI Prompts: Write C# code that uses Aspose.Cells to apply the "$#,##0.00" format to a given range, preserving other cell styles. | Show how to employ StyleFlag so only the number format is updated in Aspose.Cells for .NET. | Create a reusable method that accepts a worksheet, address string, and format pattern to apply a currency style with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, fills cells A1‑A5 with sample numbers, defines a style using the custom format "$#,##0.00", applies only the number‑format part via StyleFlag to the range, and saves the file as CurrencyNumberFormat.xlsx.
    public class ApplyCurrencyNumberFormat
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully as CurrencyNumberFormat.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a range with sample numeric values (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(1234.56 + i * 100);
            }

            // Create a style that defines a custom currency format with two decimal places
            Style currencyStyle = workbook.CreateStyle();
            // Custom format: $#,##0.00 (adjust symbol as needed)
            currencyStyle.Custom = "$#,##0.00";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Define the range to which the style will be applied (A1:A5)
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, 5, 1);
            range.ApplyStyle(currencyStyle, flag);

            // Save the workbook to a file
            workbook.Save("CurrencyNumberFormat.xlsx");
        }
    }
}
