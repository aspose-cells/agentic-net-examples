// Title: Aspose.Cells .NET: Custom number format to display negative values in red parentheses
// Description: Creates a workbook, inserts positive and negative numbers, defines a style with the custom format "#,##0.00;[Red](#,##0.00)" (positive normal, negative red and enclosed in parentheses), applies the format to a range using a StyleFlag, and saves the file as NegativeRedParentheses.xlsx.
// Keywords: Aspose.Cells | .NET | C# | custom number format | negative numbers red parentheses | Excel number format string | StyleFlag | apply style Aspose.Cells | financial Excel formatting | red negative values
// Common Searches: Aspose.Cells custom number format for negative values | display negative numbers in red parentheses using C# | how to apply only number format with StyleFlag in Aspose.Cells | Excel red parentheses format string Aspose.Cells | format negative amounts in financial reports with Aspose
// Developer Intent: Apply a custom number format that renders negative values in red parentheses to selected cells in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Financial statements where losses appear in red parentheses for quick identification. | Invoices that highlight refunds or discounts (negative amounts) with red parentheses. | Standardized accounting templates that enforce consistent negative‑value styling across multiple worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that formats a cell range so negative numbers appear in red parentheses. | Explain the components of the format string "#,##0.00;[Red](#,##0.00)" and show how to apply it via StyleFlag. | Provide a step‑by‑step tutorial for creating a reusable style that formats negative values in red parentheses and applying it to several sheets in a workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomNumberFormat
{
    // Creates a workbook, inserts positive and negative numbers, defines a style with the custom format "#,##0.00;[Red](#,##0.00)" (positive normal, negative red and enclosed in parentheses), applies the format to a range using a StyleFlag, and saves the file as NegativeRedParentheses.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                NegativeRedParenthesesDemo.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class NegativeRedParenthesesDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set sample values (positive and negative)
            sheet.Cells["A1"].PutValue(1234.56);   // Positive
            sheet.Cells["A2"].PutValue(-987.65);   // Negative

            // Create a style with a custom number format:
            // Positive numbers: normal display
            // Negative numbers: red color and enclosed in parentheses
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00;[Red](#,##0.00)";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag styleFlag = new StyleFlag();
            styleFlag.NumberFormat = true;

            // Apply the style to the range containing the sample cells
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A2");
            range.ApplyStyle(style, styleFlag);

            // Determine output path and ensure directory exists
            string outputPath = "NegativeRedParentheses.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}
