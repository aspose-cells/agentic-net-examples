// Title: Apply a green-to-yellow solid fill color to a specific cell range (B2:B10) with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to set a solid green foreground and yellow background for a range of cells | apply custom cell shading to progress column B2:B10 using Aspose.Cells .NET | Aspose.Cells style flag CellShading example for green to yellow fill | create a workbook and apply solid fill style to multiple cells in C# | set foreground and background colors of Excel cells programmatically with Aspose.Cells
// Tags: solid fill style Aspose.Cells C# | foreground background color formatting Excel Aspose.Cells | apply style to cell range B2:B10 Aspose.Cells | cell shading with StyleFlag Aspose.Cells | progress column color coding Aspose.Cells .NET

using System;
using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, defines the range B2:B10, builds a style with a green foreground and yellow background using a solid fill pattern, applies the style to the range via StyleFlag, and saves the file as ProgressGradient.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that contains progress values (e.g., B2:B10)
            Aspose.Cells.Range progressRange = sheet.Cells.CreateRange("B2:B10");

            // Create a style with a solid fill (green foreground, yellow background)
            Style gradientStyle = workbook.CreateStyle();
            gradientStyle.ForegroundColor = Color.Green;
            gradientStyle.BackgroundColor = Color.Yellow;
            gradientStyle.Pattern = BackgroundType.Solid; // solid fill

            // Apply the style to the selected range
            StyleFlag flag = new StyleFlag();
            flag.CellShading = true; // apply fill-related formatting
            progressRange.ApplyStyle(gradientStyle, flag);

            // Save the workbook with the applied style
            string outputPath = "ProgressGradient.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
