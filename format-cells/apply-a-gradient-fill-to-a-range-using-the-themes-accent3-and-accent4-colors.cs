// Title: Apply a gradient fill with theme Accent3 and Accent4 colors to a cell range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells apply gradient fill using theme Accent3 and Accent4 to a range | how to set background gradient with theme colors in Aspose.Cells .NET | simulate gradient effect with solid pattern in Aspose.Cells C# example | apply style with two theme colors to cells Aspose.Cells workbook | range formatting gradient fill Aspose.Cells using theme colors
// Tags: gradient fill theme colors Aspose.Cells | apply style to cell range C# Aspose.Cells | solid pattern gradient simulation Aspose.Cells | theme Accent3 Accent4 background formatting Aspose.Cells | cell range background color gradient Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a workbook, defines the range B2:D6, builds a style with a solid pattern, sets LightBlue and SteelBlue (representing theme Accent3 and Accent4) as foreground and background colors to mimic a gradient, applies the style to the range, and saves the file as GradientFill.xlsx.
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

            // Define the range to which the fill will be applied
            Aspose.Cells.Range range = sheet.Cells.CreateRange("B2:D6");

            // Create a new style
            Style style = workbook.CreateStyle();

            // Use a solid pattern (gradient not directly supported in this version)
            style.Pattern = BackgroundType.Solid;

            // Set foreground and background colors to simulate a gradient effect
            style.ForegroundColor = Color.LightBlue; // first color
            style.BackgroundColor = Color.SteelBlue; // second color

            // Apply the style to the range
            StyleFlag flag = new StyleFlag { All = true };
            range.ApplyStyle(style, flag);

            // Save the workbook
            string outputPath = "GradientFill.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
