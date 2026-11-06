// Title: How to apply a custom number format in Aspose.Cells for .NET (C#) that shows zero values as a dash
// Common Searches: Aspose.Cells C# custom number format to display zero as a dash | How to hide zero values with a hyphen in Excel using Aspose.Cells .NET | Set positive, negative, and zero patterns in Aspose.Cells style | Apply custom numeric format to a range of cells in Aspose.Cells C# | Create workbook with custom number format string in Aspose.Cells
// Tags: custom numeric format pattern dash zero Aspose.Cells | apply style with custom number format C# | set number format for cell range Aspose.Cells | Excel .xlsx hide zero values using Aspose.Cells | C# Aspose.Cells style flag numberformat

using Aspose.Cells;
using System;
using System.IO;

// The example creates a Workbook, writes positive, negative, and zero values to cells A1‑A3, defines the custom format "#,##0;-#,##0;-" to render zeros as a dash, builds a Style with this format, applies it to the range via a StyleFlag, and saves the file as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: positive, negative, and zero values
            sheet.Cells["A1"].PutValue(123);
            sheet.Cells["A2"].PutValue(-456);
            sheet.Cells["A3"].PutValue(0);

            // Custom number format:
            //   Positive numbers:  #,##0
            //   Negative numbers: -#,##0
            //   Zero values:      -
            string customFormat = "#,##0;-#,##0;-";

            // Create a style with the custom format
            Style style = workbook.CreateStyle();
            style.Custom = customFormat;

            // Indicate that the number format should be applied
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the range A1:A3
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A3");
            range.ApplyStyle(style, flag);

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomNumberFormat.xlsx");

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
