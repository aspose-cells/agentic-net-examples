// Title: How to style a header row and freeze the first row in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells apply bold gray header style to first row and freeze panes | Aspose.Cells .NET freeze top row after formatting header cells | Set background color and center alignment for A1:J1 with Aspose.Cells and keep row visible while scrolling | Create styled header and freeze first worksheet row using Aspose.Cells C# example | Aspose.Cells FreezePanes usage after applying style to header range
// Tags: header row styling with Aspose.Cells | freeze top worksheet row using FreezePanes | bold gray header background C# Aspose.Cells | style cell range A1:J1 Aspose.Cells | export workbook to xlsx with Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, defines a bold, centered style with a light‑gray solid background, applies the style to cells A1 through J1 on the first worksheet, freezes the first row using FreezePanes, ensures the output directory exists, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Define a style for the header row
            var headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;

            // Apply the style to the header row (A1 to J1)
            var styleFlag = new StyleFlag { All = true };
            var headerRange = sheet.Cells.CreateRange("A1:J1");
            headerRange.ApplyStyle(headerStyle, styleFlag);

            // Freeze the first row so it stays visible while scrolling
            // Parameters: totalRows, totalColumns, frozenRows, frozenColumns
            sheet.FreezePanes(0, 0, 1, 0);

            // Determine output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
