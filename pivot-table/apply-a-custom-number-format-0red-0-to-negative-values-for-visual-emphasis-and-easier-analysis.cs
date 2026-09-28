// Title: Apply a custom number format '#,##0;[Red]-#,##0' to a cell range with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style using the custom format '#,##0;[Red]-#,##0' and applies it to cells A1 through A3 with Aspose.Cells. | Show how to use a StyleFlag to modify only the number‑format property of a range in an Aspose.Cells workbook. | Produce an .xlsx file where negative numbers appear in red by applying a custom number format via Aspose.Cells for .NET.
// Common Searches: aspnet how to format negative numbers in red using Aspose.Cells | c# apply custom number format to a range of cells with Aspose.Cells | using StyleFlag to set only number format in Aspose.Cells example | create workbook with '#,##0;[Red]-#,##0' format in C# | how to use range.ApplyStyle for custom number formatting in Aspose.Cells
// Tags: custom number format negative values Aspose.Cells | StyleFlag number format only C# | highlight negative numbers red Excel Aspose.Cells | range.ApplyStyle custom number format .NET | create workbook with custom format Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a new workbook, fills cells A1‑A3 with sample values, defines a style with the custom number format '#,##0;[Red]-#,##0', uses a StyleFlag to apply only the number‑format attribute, applies the style to the range A1:A3, and saves the file as CustomNumberFormat.xlsx.
class ApplyCustomNumberFormat
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with sample data, including negative values
            sheet.Cells["A1"].PutValue(12345);
            sheet.Cells["A2"].PutValue(-6789);
            sheet.Cells["A3"].PutValue(0);

            // Create a style with the custom number format for negative values
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0;[Red]-#,##0";

            // Use a StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the desired range
            AsposeRange range = sheet.Cells.CreateRange("A1:A3");
            range.ApplyStyle(customStyle, flag);

            // Determine output path and save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomNumberFormat.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
