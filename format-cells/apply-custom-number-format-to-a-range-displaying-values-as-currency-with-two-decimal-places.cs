// Title: Apply Euro Currency Custom Number Format (2 Decimals) to a Cell Range with Aspose.Cells for .NET
// Description: Creates a workbook, inserts numeric values, defines a custom Euro format "_-€ #,##0.00;[Red]_-€ -#,##0.00", uses StyleFlag to apply only the number‑format, formats the range A1:A2, and saves the file as CurrencyFormat.xlsx.
// Keywords: Aspose.Cells | C# | .NET | custom number format | Euro currency format | two decimal places | StyleFlag | apply format to range | Excel currency formatting | save workbook
// Common Searches: Aspose.Cells set custom Euro format C# | apply number format to a range using StyleFlag Aspose.Cells | C# format cells as currency with two decimals Aspose | how to use custom number format in Aspose.Cells | save Excel file with currency formatting Aspose.Cells
// Developer Intent: The developer needs to display numeric values as Euro currency with two decimal places for a specific range without altering other cell styles.
// Use Cases: Standardize monetary columns in financial reports to Euro format while preserving existing fonts and borders. | Highlight negative amounts in red within an invoice sheet using a single custom number format. | Prepare exported spreadsheets for accounting systems by applying consistent currency formatting to all monetary cells.
// AI Prompts: Generate C# code with Aspose.Cells that formats cells B2:B20 as US dollars with two decimal places. | Explain how StyleFlag can isolate number‑format changes from other style attributes in Aspose.Cells. | Create a custom number format in Aspose.Cells that shows negative values in red parentheses.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts numeric values, defines a custom Euro format "_-€ #,##0.00;[Red]_-€ -#,##0.00", uses StyleFlag to apply only the number‑format, formats the range A1:A2, and saves the file as CurrencyFormat.xlsx.
    public class ApplyCurrencyFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some cells with numeric values
                sheet.Cells["A1"].PutValue(1234.56);
                sheet.Cells["A2"].PutValue(7890.12);

                // Create a style with a custom currency format (two decimal places)
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Custom = "_-€ #,##0.00;[Red]_-€ -#,##0.00";

                // Use StyleFlag to apply only the number format part of the style
                StyleFlag styleFlag = new StyleFlag();
                styleFlag.NumberFormat = true;

                // Define the range to which the style will be applied (A1:A2)
                Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, 2, 1);
                range.ApplyStyle(currencyStyle, styleFlag);

                // Save the workbook to a file
                string outputPath = "CurrencyFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCurrencyFormat.Run();
        }
    }
}
