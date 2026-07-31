// Title: Aspose.Cells for .NET – Apply Custom Number Format to Show Negative Values in Red Parentheses
// Description: Creates a workbook, writes a positive and a negative value to cells A1 and A2, defines a style with the custom format "#,##0.00;[Red](#,##0.00)" (positive numbers normal, negatives red and enclosed in parentheses), applies the format to the range via a StyleFlag, and saves the file as NegativeNumberCustomFormat.xlsx.
// Keywords: Aspose.Cells | .NET | custom number format | negative numbers red parentheses | StyleFlag | Excel formatting | C# Aspose.Cells example | financial spreadsheet formatting | Excel custom format string | range styling
// Common Searches: Aspose.Cells format negative numbers red parentheses | C# custom number format for negative values in Excel | How to use StyleFlag to apply number format in Aspose.Cells | Set red parentheses format for negatives with Aspose.Cells .NET | Apply custom number format to a cell range using Aspose.Cells
// Developer Intent: The developer wants to display negative numbers in red parentheses while keeping positive numbers in the default numeric format.
// Use Cases: Financial statements where debit amounts appear in red parentheses for quick identification. | Invoices that show refunds as red numbers inside parentheses, distinguishing them from regular charges. | Exported accounting reports that automatically apply currency symbols and red parentheses to negative balances across a column.
// AI Prompts: Show how to add a currency symbol to the custom format so positives display "$#,##0.00" and negatives show red parentheses with the symbol. | Provide code to apply the red‑parentheses format to an entire column (e.g., column B) using Aspose.Cells. | Explain how to combine a font‑color style flag with the number‑format flag to set both font color and custom number format in a single operation.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a workbook, writes a positive and a negative value to cells A1 and A2, defines a style with the custom format "#,##0.00;[Red](#,##0.00)" (positive numbers normal, negatives red and enclosed in parentheses), applies the format to the range via a StyleFlag, and saves the file as NegativeNumberCustomFormat.xlsx.
    public class NegativeNumberCustomFormat
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
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

            // Insert sample values: one positive and one negative
            sheet.Cells["A1"].PutValue(1234.56);
            sheet.Cells["A2"].PutValue(-1234.56);

            // Create a style with a custom number format.
            // Positive numbers: 1,234.56
            // Negative numbers: red color and enclosed in parentheses.
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00;[Red](#,##0.00)";

            // Use StyleFlag to apply only the number format part of the style.
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the range A1:A2.
            AsposeRange range = sheet.Cells.CreateRange("A1", "A2");
            range.ApplyStyle(style, flag);

            // Save the workbook to a file.
            workbook.Save("NegativeNumberCustomFormat.xlsx");
        }
    }
}
