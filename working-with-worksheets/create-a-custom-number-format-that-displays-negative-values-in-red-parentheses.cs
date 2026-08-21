// Title: Aspose.Cells .NET – Custom number format to display negatives in red parentheses
// Description: Shows how to build a Workbook, write positive and negative values, define the format "#,##0.00;[Red](#,##0.00)" and apply it with a StyleFlag so only the number‑format changes, then save the spreadsheet.
// Keywords: Aspose.Cells | custom number format | red parentheses | negative numbers | StyleFlag | C# | .NET | Excel formatting | financial reporting | cell style
// Common Searches: Aspose.Cells format negative numbers red parentheses | C# custom number format string for Excel | Apply only number format with StyleFlag Aspose.Cells | How to show losses in red brackets using Aspose.Cells | Excel custom format #,##0.00;[Red](#,##0.00) Aspose
// Developer Intent: Create and apply a custom number format that renders negative values in red parentheses without altering other cell style attributes.
// Use Cases: Financial statements where deficits appear in red brackets for quick visual identification. | Automated Excel reports that need consistent monetary formatting while highlighting negative balances. | Styling a column of figures where only the numeric representation changes, preserving existing fonts and borders.
// AI Prompts: Write C# code with Aspose.Cells to define a custom number format that shows negatives in red parentheses and apply it to a range. | Explain how to modify the format to add a currency symbol and switch the negative color to blue in Aspose.Cells. | Provide step‑by‑step instructions for using StyleFlag to change only the number‑format of selected cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCustomNumberFormat
{
    // Shows how to build a Workbook, write positive and negative values, define the format "#,##0.00;[Red](#,##0.00)" and apply it with a StyleFlag so only the number‑format changes, then save the spreadsheet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set sample values (positive and negative) in cells A1 and A2
                sheet.Cells["A1"].PutValue(1234.56);   // Positive number
                sheet.Cells["A2"].PutValue(-1234.56);  // Negative number

                // Create a style with a custom number format:
                // Positive numbers: normal format with two decimals
                // Negative numbers: red color and enclosed in parentheses
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "#,##0.00;[Red](#,##0.00)";

                // Use StyleFlag to apply only the number format part of the style
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Apply the style to the range A1:A2
                AsposeRange range = sheet.Cells.CreateRange("A1", "A2");
                range.ApplyStyle(customStyle, flag);

                // Save the workbook to verify the formatting
                workbook.Save("CustomNumberFormat_Output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
