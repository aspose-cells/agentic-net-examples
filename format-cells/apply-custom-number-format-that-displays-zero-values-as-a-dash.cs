// Title: How to use Aspose.Cells for .NET to apply a custom number format that shows zero values as a dash
// AI Prompts: Generate C# code that creates a workbook, defines the custom format "0;-0;\"-\"" and applies it to a cell range using Aspose.Cells. | Write a method that sets a style with a custom numeric format to display zeros as a dash while preserving positive and negative patterns. | Demonstrate saving an Excel file after applying a dash‑for‑zero format to a specific column with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# custom number format for zero values dash | Set Excel cell format to display '-' for zero using Aspose.Cells | C# Aspose.Cells apply 0;-0;"-" format to a range | How to format zero as a dash in an Excel workbook with Aspose.Cells .NET
// Tags: Aspose.Cells custom numeric style | apply number format to cell range C# | display dash for zero in Excel | Aspose.Cells style flag numberformat usage | save formatted workbook .NET

using Aspose.Cells;
using System;

// Creates a workbook, inserts sample data, defines a custom number format (positive: 0, negative: -0, zero: "-"), applies the style to cells A1:A4, and saves the file as CustomNumberFormat.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data: includes positive, negative and zero values
        sheet.Cells["A1"].PutValue(123);
        sheet.Cells["A2"].PutValue(-45);
        sheet.Cells["A3"].PutValue(0);
        sheet.Cells["A4"].PutValue(78.9);

        // Define a custom number format:
        //   Positive numbers: 0
        //   Negative numbers: -0
        //   Zero values: display a dash ("-")
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "0;-0;\"-\"";

        // Apply the custom style to the target range
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // only affect number format
        sheet.Cells.CreateRange("A1:A4").ApplyStyle(customStyle, flag);

        // Save the workbook to a file
        workbook.Save("CustomNumberFormat.xlsx");
    }
}
