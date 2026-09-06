// Title: How to format negative numbers in red parentheses using a custom number format in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Aspose.Cells style with the custom number format "#,##0;[Red](#,##0)" and applies it to a target cell range. | Show how to save an Aspose.Cells workbook after applying a style that renders negative values in red parentheses.
// Common Searches: Aspose.Cells C# custom number format to display negative values in red parentheses | apply red parentheses number format to a range using Aspose.Cells .NET | how to set negative number display style in an Excel file with Aspose.Cells | C# Aspose.Cells style for negative numbers red and enclosed in parentheses | save workbook after formatting negative numbers with Aspose.Cells
// Tags: Aspose.Cells red parentheses style | apply style to cell range C# Aspose.Cells | save workbook after formatting negatives Aspose.Cells | C# set number format pattern Aspose.Cells | Excel negative value formatting Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, writes positive, negative, and zero values to cells A1‑A3, defines a style with the custom format "#,##0;[Red](#,##0)" to show negative numbers in red parentheses, applies the style to the range A1:A3, and saves the file as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: positive, negative, and zero values
            sheet.Cells["A1"].PutValue(1234);
            sheet.Cells["A2"].PutValue(-5678);
            sheet.Cells["A3"].PutValue(0);

            // Define a custom number format:
            // Positive numbers: normal display
            // Negative numbers: red color and enclosed in parentheses
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0;[Red](#,##0)";

            // Apply the custom style to the target range
            AsposeRange range = sheet.Cells.CreateRange("A1:A3");
            range.SetStyle(customStyle); // Apply the style (including number format)

            // Define output file path
            string outputPath = "CustomNumberFormat.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
