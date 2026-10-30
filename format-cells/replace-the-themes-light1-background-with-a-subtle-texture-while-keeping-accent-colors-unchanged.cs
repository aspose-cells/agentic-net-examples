// Title: Apply a solid background fill to every cell of the first worksheet while keeping Excel theme accent colors unchanged using Aspose.Cells for .NET (C#)
// Common Searches: how to apply a solid fill to all cells in a worksheet with Aspose.Cells C# | preserve Excel theme accent colors when changing worksheet background using Aspose.Cells | bulk style application to entire sheet without affecting theme colors Aspose.Cells .NET | replace Light1 theme background with custom fill in C# Aspose.Cells example
// Tags: Aspose.Cells solid fill entire worksheet | preserve theme accent colors Aspose.Cells | C# bulk cell styling StyleFlag All | use workbook default foreground color as fill Aspose.Cells | replace Light1 background with custom fill Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, creates a solid‑fill style using the workbook's default foreground color, applies this style to every cell in the first worksheet via a StyleFlag, and saves the modified file while leaving the theme's accent colors untouched.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Use a default foreground color (e.g., the workbook's default style color)
            Color foregroundColor = workbook.DefaultStyle.ForegroundColor.IsEmpty
                ? Color.Black
                : workbook.DefaultStyle.ForegroundColor;

            // Create a new style with a solid background pattern
            Style textureStyle = workbook.CreateStyle();
            textureStyle.Pattern = BackgroundType.Solid;
            textureStyle.ForegroundColor = foregroundColor;
            textureStyle.BackgroundColor = Color.White;

            // Apply the style to the entire first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            StyleFlag flag = new StyleFlag { All = true };
            sheet.Cells.ApplyStyle(textureStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
