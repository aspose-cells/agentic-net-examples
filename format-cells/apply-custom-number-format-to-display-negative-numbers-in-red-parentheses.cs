// Title: Apply a custom number format to show negative values in red parentheses with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# custom number format red parentheses for negative values | format negative numbers in red parentheses using Aspose.Cells .NET | apply custom number format to a cell range in Aspose.Cells C# example | C# Aspose.Cells style for red parentheses negative numbers | how to use custom number format string "#,##0.00;[Red](#,##0.00)" in Aspose.Cells
// Tags: Aspose.Cells custom number format styling | red parentheses negative number format | C# apply style to cell range Aspose.Cells | Excel number format string Aspose.Cells .NET | style.Custom property usage Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, writes a positive and a negative value, defines the custom format "#,##0.00;[Red](#,##0.00)" to display negatives in red parentheses, applies the style to cells A1:A2, and saves the file as FormattedNumbers.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Sample values: positive and negative
        worksheet.Cells["A1"].PutValue(1234.56);
        worksheet.Cells["A2"].PutValue(-1234.56);

        // Custom number format: positive numbers normal, negative numbers in red parentheses
        string customFormat = "#,##0.00;[Red](#,##0.00)";

        // Create a style with the custom format
        Style style = workbook.CreateStyle();
        style.Custom = customFormat;

        // Apply the style to the target range
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;
        worksheet.Cells.CreateRange("A1:A2").ApplyStyle(style, flag);

        // Save the workbook
        workbook.Save("FormattedNumbers.xlsx");
    }
}
