// Title: Create a new workbook, merge cells L5:M7, and apply an italic gray style with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge cells L5:M7 and set italic font with gray background | How to apply a custom style to a merged range using Aspose.Cells .NET | Create workbook and style merged cells with solid gray fill in C# | Save an Excel file after merging cells and applying italic text with Aspose.Cells | Example of merging a rectangular range and formatting it with Aspose.Cells for .NET
// Tags: merge cells range Aspose.Cells C# | apply custom style Aspose.Cells | italic font gray fill Aspose.Cells | save workbook as xlsx Aspose.Cells | create workbook Aspose.Cells .NET

using System;
using Aspose.Cells;
using System.Drawing;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example demonstrates how to instantiate a new Workbook, merge the cell range L5:M7 on the first worksheet, define a custom Style with italic font and a solid gray fill, apply this style to the merged range, and finally save the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells L5:M7 (range "L5:M7")
            // Row index 4 (zero‑based for L5), column index 11 (L), 3 rows, 2 columns
            sheet.Cells.Merge(4, 11, 3, 2);

            // Create a custom style: italic font and gray fill
            Style style = workbook.CreateStyle();
            style.Font.IsItalic = true;
            style.ForegroundColor = Color.Gray;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the merged range
            StyleFlag styleFlag = new StyleFlag { All = true };
            AsposeRange mergedRange = sheet.Cells.CreateRange("L5:M7");
            mergedRange.ApplyStyle(style, styleFlag);

            // Save the workbook to a file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
