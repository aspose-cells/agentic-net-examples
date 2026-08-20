// Title: Apply a custom number format with thousand separators and parentheses for negatives using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, writes a positive and a negative value, defines the custom format "#,##0.00;(#,##0.00)" to add thousand separators, show two decimals and wrap negative numbers in parentheses, applies the format to cells A1:A2 with a StyleFlag that targets only the number format, and saves the file as CustomNumberFormat.xlsx.
// Keywords: Aspose.Cells custom numeric format | C# thousand separator format | negative numbers parentheses Excel | StyleFlag number format only | Excel custom format Aspose .NET
// Common Searches: Aspose.Cells format negative numbers with parentheses | C# custom number format with commas in Aspose.Cells | How to use StyleFlag to set only number format in Aspose.Cells | Create Excel custom numeric format string using Aspose.Cells
// Developer Intent: Generate and apply a custom numeric style that inserts commas for thousands and encloses negative values in parentheses without altering other cell attributes.
// Use Cases: Financial reports where negatives appear in parentheses while positives use standard comma separators. | Applying a consistent monetary format to an entire column of generated spreadsheets. | Automating Excel exports that require precise number formatting for both positive and negative amounts.
// AI Prompts: Write C# code with Aspose.Cells that applies the format '#,##0.00;(#,##0.00)' to a given range while preserving existing cell styles. | Explain the role of StyleFlag in Aspose.Cells when you want to modify only the number format of a style. | Provide a reusable method in C# that formats any worksheet column with thousand separators and parentheses for negative values using Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This example creates a workbook, writes a positive and a negative value, defines the custom format "#,##0.00;(#,##0.00)" to add thousand separators, show two decimals and wrap negative numbers in parentheses, applies the format to cells A1:A2 with a StyleFlag that targets only the number format, and saves the file as CustomNumberFormat.xlsx.
class CustomNumberFormatDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a positive and a negative number in cells A1 and A2
            sheet.Cells["A1"].PutValue(1234567.89);
            sheet.Cells["A2"].PutValue(-1234567.89);

            // Create a style with a custom number format:
            // "#,##0.00;(#,##0.00)" adds thousand separators,
            // shows two decimal places, and encloses negative numbers in parentheses
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00;(#,##0.00)";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the range A1:A2
            AsposeRange range = sheet.Cells.CreateRange("A1", "A2");
            range.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("CustomNumberFormat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
