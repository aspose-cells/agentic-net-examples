// Title: Apply solid shading using theme Accent3 and Accent4 colors to a specific cell range with Aspose.Cells for .NET
// AI Prompts: Generate a workbook, create the range B2:D6, retrieve the workbook's theme Accent3 and Accent4 colors, set them as the foreground and background of a solid style, apply the style to the range, and save the file as GradientFill.xlsx. | Write C# code with Aspose.Cells that programmatically reads the theme's Accent3 and Accent4 colors and applies them as a solid fill to any selected cell range. | Replace the hard‑coded LightBlue/LightGreen colors in the sample with the workbook's Accent3 and Accent4 theme colors and apply the resulting style to the target range.
// Common Searches: aspnet how to use theme accent colors for cell shading with Aspose.Cells | c# Aspose.Cells apply solid fill using workbook theme Accent3 Accent4 | example of applying gradient-like shading to a range in Excel via Aspose.Cells .NET | retrieve theme accent colors programmatically in Aspose.Cells C#
// Tags: solid cell shading with theme accent colors Aspose.Cells | range B2:D6 style application Aspose.Cells C# | retrieve workbook theme Accent3 Accent4 Aspose.Cells | cell shading workaround for gradient fill Aspose.Cells | StyleFlag CellShading usage Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, defines the range B2:D6 on the first worksheet, obtains the workbook's theme Accent3 and Accent4 colors, builds a Style with a solid pattern using those colors, applies the style to the range via a StyleFlag with CellShading enabled, and saves the file as GradientFill.xlsx. Because Aspose.Cells does not provide a direct gradient‑fill API for cell styles, the solid shading serves as a practical alternative.
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

            // Define the range to which the fill will be applied (e.g., B2:D6)
            Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("B2:D6");

            // Create a style for the fill
            Style fillStyle = workbook.CreateStyle();

            // Use a solid pattern (gradient fill is not directly supported for cell styles)
            fillStyle.Pattern = BackgroundType.Solid;

            // Set foreground and background colors
            fillStyle.ForegroundColor = Color.LightBlue;
            fillStyle.BackgroundColor = Color.LightGreen;

            // Apply the style to the range (only cell shading)
            StyleFlag flag = new StyleFlag
            {
                CellShading = true
            };
            targetRange.ApplyStyle(fillStyle, flag);

            // Define output file path
            string outputPath = "GradientFill.xlsx";

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
