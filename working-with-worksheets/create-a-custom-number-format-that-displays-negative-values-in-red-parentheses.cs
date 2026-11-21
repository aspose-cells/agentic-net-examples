// Title: Apply a custom number format in Aspose.Cells for .NET to show negative values in red parentheses
// Common Searches: Aspose.Cells C# custom number format red parentheses for negative numbers | how to set negative numbers red and in parentheses using Aspose.Cells | C# Aspose.Cells style custom format string for red negative values | formatting Excel cells with red parentheses for negatives in Aspose.Cells .NET | apply custom numeric format '#,##0.00_);[Red](#,##0.00)' with Aspose.Cells
// Tags: custom numeric format red parentheses Aspose.Cells | negative number styling C# Aspose.Cells | Style.Custom property Excel format Aspose.Cells | apply custom number format workbook Aspose.Cells | Excel cell formatting red negative values Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, defines a custom number format '#,##0.00_);[Red](#,##0.00)' that displays negative numbers in red parentheses, applies the style to cells A1 and A2, and saves the workbook as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert sample values: positive and negative
        sheet.Cells["A1"].PutValue(1234.56);
        sheet.Cells["A2"].PutValue(-1234.56);

        // Custom number format:
        // Positive numbers: normal display
        // Negative numbers: red color and enclosed in parentheses
        string customFormat = "#,##0.00_);[Red](#,##0.00)";

        // Create a style and assign the custom format
        Style style = workbook.CreateStyle();
        style.Custom = customFormat;

        // Apply the style to the cells
        sheet.Cells["A1"].SetStyle(style);
        sheet.Cells["A2"].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("CustomNumberFormat.xlsx");
    }
}
