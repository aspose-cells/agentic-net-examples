// Title: Aspose.Cells .NET – Apply Custom Number Format with Thousand Separator and Parentheses for Negatives
// Description: Demonstrates how to create a workbook, insert a numeric value, and apply a custom format "#,##0.00;(#,##0.00)" that adds commas, forces two decimal places, and encloses negative numbers in parentheses. The example uses a Style with a StyleFlag to modify only the NumberFormat property and saves the result as an XLSX file. Suitable for global .NET developers working with Excel reports.
// Keywords: Aspose.Cells custom number format | C# thousand separator | negative numbers parentheses | StyleFlag NumberFormat | Excel number formatting .NET | financial reporting Excel | Aspose.Cells example
// Common Searches: Aspose.Cells format numbers with commas C# | Show negative values in parentheses using Aspose.Cells | Apply custom number format to a single cell Aspose.Cells .NET | StyleFlag only number format Aspose.Cells example | C# Excel custom format #,##0.00;(#,##0.00)
// Developer Intent: Create a workbook and format a cell so large numbers use commas and negative values appear in parentheses, without altering other cell styles.
// Use Cases: Financial statements where millions need comma separators and losses are shown in parentheses. | Invoices that require two‑decimal monetary values with clear negative‑amount notation. | Automated data exports that enforce a consistent numeric style across multiple worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that applies the custom format "#,##0.00;(#,##0.00)" to range A1 while preserving existing fonts and borders. | Explain how to use StyleFlag in Aspose.Cells to change only the NumberFormat of a cell without affecting other style attributes. | Show a step‑by‑step example of creating a workbook, inserting a value, defining a custom number format with thousand separators and parentheses, and saving the file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNumberFormatDemo
{
    // Demonstrates how to create a workbook, insert a numeric value, and apply a custom format "#,##0.00;(#,##0.00)" that adds commas, forces two decimal places, and encloses negative numbers in parentheses. The example uses a Style with a StyleFlag to modify only the NumberFormat property and saves the result as an XLSX file. Suitable for global .NET developers working with Excel reports.
    public class ApplyCustomNumberFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set a numeric value in cell A1
                sheet.Cells["A1"].PutValue(1234567.89);

                // Create a style and define a custom number format:
                // - Thousand separator (",")
                // - Two decimal places
                // - Negative numbers displayed in parentheses
                Style style = workbook.CreateStyle();
                style.Custom = "#,##0.00;(#,##0.00)";

                // Use StyleFlag to apply only the number format part of the style
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Apply the style to the target cell (range of one cell)
                Aspose.Cells.Range range = sheet.Cells.CreateRange("A1");
                range.ApplyStyle(style, flag);

                // Save the workbook to verify the formatting
                string outputPath = "CustomNumberFormatDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            ApplyCustomNumberFormat.Run();
        }
    }
}
