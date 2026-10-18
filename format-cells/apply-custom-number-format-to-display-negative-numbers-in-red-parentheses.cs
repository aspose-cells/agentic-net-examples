// Title: How to format negative numbers in red parentheses using a custom number format with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# custom number format red parentheses for negative values | apply red parentheses style to negative numbers in Excel workbook using Aspose.Cells | C# code to display negative numbers in red and parentheses with Aspose.Cells | set number format '#,##0;[Red](#,##0)' on a cell range using Aspose.Cells .NET | formatting negative currency values in red parentheses with Aspose.Cells API
// Tags: Aspose.Cells custom number format syntax | StyleFlag number format application Aspose.Cells | range.ApplyStyle usage Aspose.Cells | Excel negative number red parentheses formatting .NET | Workbook.Save with formatted cells Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, writes a positive and a negative value to cells A1 and A2, defines a custom style with the format "#,##0;[Red](#,##0)" to show negative numbers in red parentheses, applies the style to the A1:A2 range using a StyleFlag that targets only the number format, and saves the file as FormattedNumbers.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert sample numbers: a positive and a negative value
            sheet.Cells["A1"].PutValue(123);
            sheet.Cells["A2"].PutValue(-456);

            // Create a style with a custom number format:
            // Positive numbers: normal display
            // Negative numbers: red color and enclosed in parentheses
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0;[Red](#,##0)";

            // Define which style attributes to apply (only number format)
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the custom style to the range containing the numbers
            AsposeRange range = sheet.Cells.CreateRange("A1:A2");
            range.ApplyStyle(customStyle, flag);

            // Determine output file path
            string outputPath = "FormattedNumbers.xlsx";

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
