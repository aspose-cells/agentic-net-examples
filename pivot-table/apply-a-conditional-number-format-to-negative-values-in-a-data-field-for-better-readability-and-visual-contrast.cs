// Title: C# – Apply Conditional Number Format to Highlight Negative Values in Excel with Aspose.Cells
// Description: This example creates a new workbook, writes positive, negative and zero values to cells A1‑A5, defines a custom number format "#,##0;[Red]-#,##0;0" that shows negatives in red, applies the format to the range using a StyleFlag that targets only the number format, and saves the file as an .xlsx workbook. The same technique can be used for regular sheets or pivot tables.
// Keywords: Aspose.Cells | C# | .NET | conditional number format | negative numbers red | custom Excel number format | StyleFlag | apply style to range | Excel formatting example | pivot table formatting
// Common Searches: How to format negative numbers in red with Aspose.Cells C# | Aspose.Cells custom number format positive negative zero | Apply number format to a cell range using Aspose.Cells .NET | Conditional formatting for negative values in Excel via code | Aspose.Cells example for number format in pivot tables
// Developer Intent: Apply a custom number format that displays negative values in red (or another style) to a specific cell range in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Financial reports where losses appear in red for instant visual cue. | Profit‑and‑loss statements that need distinct formatting for positive, negative, and zero balances. | Sales dashboards that highlight negative adjustments while keeping regular figures unchanged.
// AI Prompts: Generate C# code with Aspose.Cells to format negative numbers in blue parentheses. | Show how to set different number formats for positive, negative, and zero values inside a pivot table using Aspose.Cells. | Explain how to add a currency symbol and thousand separators to the custom format while keeping negatives red.

using System;
using Aspose.Cells;
using System.Drawing;

// This example creates a new workbook, writes positive, negative and zero values to cells A1‑A5, defines a custom number format "#,##0;[Red]-#,##0;0" that shows negatives in red, applies the format to the range using a StyleFlag that targets only the number format, and saves the file as an .xlsx workbook. The same technique can be used for regular sheets or pivot tables.
class ConditionalNegativeNumberFormat
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with both positive and negative numbers
        sheet.Cells["A1"].PutValue(1500);
        sheet.Cells["A2"].PutValue(-750);
        sheet.Cells["A3"].PutValue(3000);
        sheet.Cells["A4"].PutValue(-1200);
        sheet.Cells["A5"].PutValue(0);

        // Create a custom number format:
        //   Positive numbers:  #,##0
        //   Negative numbers:  [Red]-#,##0  (displayed in red with minus sign)
        //   Zero:               0
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "#,##0;[Red]-#,##0;0";

        // Configure a StyleFlag to apply only the number format
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the style to the range containing the data
        // Fully qualify the Range type as required
        Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A5");
        range.ApplyStyle(customStyle, flag);

        // Save the workbook
        workbook.Save("ConditionalNegativeNumberFormat.xlsx");
    }
}
