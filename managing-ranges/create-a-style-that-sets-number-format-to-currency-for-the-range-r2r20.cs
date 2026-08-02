// Title: C# Aspose.Cells Example: Apply Built‑In Currency Format to Range R2:R20
// Description: Shows how to create a Style with the built‑in currency number format (Number = 5), define the R2:R20 range, apply the style via SetStyle, add sample values, and save the workbook as CurrencyStyleRange.xlsx.
// Keywords: Aspose.Cells C# | currency number format | Style SetStyle | CreateStyle | CreateRange R2:R20 | Excel currency formatting | built‑in format 5 | financial cells styling | range styling Aspose.Cells | C# Excel automation
// Common Searches: Aspose.Cells set currency format for R2 to R20 C# | apply style to specific column range Aspose.Cells | C# example of currency style on Excel range | how to use built‑in number format 5 with Aspose.Cells | format financial column as currency using Aspose.Cells
// Developer Intent: Create a currency style and apply it to cells R2‑R20 in an Excel workbook.
// Use Cases: Display monetary values in a single column with consistent USD formatting. | Prepare financial reports where a predefined currency style is reused across multiple sheets. | Generate test workbooks that showcase currency formatting for validation pipelines.
// AI Prompts: Generate C# code with Aspose.Cells that applies a built‑in currency format to the range A1:A15 and saves the file. | Explain how to create a reusable currency style and assign it to several non‑contiguous ranges in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Shows how to create a Style with the built‑in currency number format (Number = 5), define the R2:R20 range, apply the style via SetStyle, add sample values, and save the workbook as CurrencyStyleRange.xlsx.
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

            // Define the range R2:R20 using the Aspose.Cells.Range alias to avoid conflict with System.Range
            AsposeRange range = worksheet.Cells.CreateRange("R2", "R20");

            // Apply the currency style to the entire range
            range.SetStyle(currencyStyle);

            // Fill the range with sample numeric values to see the formatting
            for (int row = 1; row <= 19; row++) // rows are zero‑based; row 1 = R2, row 19 = R20
            {
                worksheet.Cells[row, 17].PutValue(1234.56 + row); // column 17 = column R
            }

            // Save the workbook
            workbook.Save("CurrencyStyleRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
