// Title: How to apply a red-colored conditional number format to negative values in an Excel column using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a custom style with the format '#,##0;[Red]-#,##0' and applies it to a range of cells using Aspose.Cells. | Show how to use StyleFlag to apply only the number‑format attribute of a style to a specific column in a workbook. | Write a complete Aspose.Cells example that inserts sample positive and negative numbers, formats negatives in red, and saves the file.
// Common Searches: Aspose.Cells C# format negative numbers in red within a column | apply custom number format to Excel range using StyleFlag Aspose.Cells | conditional number formatting for negative values in .NET workbook | example code for red negative number style in Aspose.Cells | how to use CreateStyle and ApplyStyle for conditional formatting in C#
// Tags: negative-number red style Aspose.Cells .NET | StyleFlag numberformat usage Aspose.Cells | custom format pattern '#,##0;[Red]-#,##0' Aspose | apply style to cell range Aspose.Cells C# | conditional number formatting Excel Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, adds a header and mixed positive/negative values to column A, defines a custom number format '#,##0;[Red]-#,##0' that displays negative numbers in red, uses a StyleFlag to apply only the number‑format part of the style to the range A2:A5, and saves the result as ConditionalNumberFormatDemo.xlsx.
class ConditionalNumberFormatDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with both positive and negative numbers
            sheet.Cells["A1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue(1500);
            sheet.Cells["A3"].PutValue(-750);
            sheet.Cells["A4"].PutValue(300);
            sheet.Cells["A5"].PutValue(-1200);

            // Define a custom number format: positive numbers normal, negative numbers in red
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0;[Red]-#,##0";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the data range (excluding the header)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A2:A5");
            range.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("ConditionalNumberFormatDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
