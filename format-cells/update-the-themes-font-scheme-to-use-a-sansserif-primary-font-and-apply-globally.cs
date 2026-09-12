// Title: Set a sans‑serif default font for an entire Excel workbook and apply it to all existing cells using Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook with Aspose.Cells, change its DefaultStyle.Font to "Arial", and propagate the font to every used range in all worksheets. | Write a reusable C# method that takes a font name, updates the workbook’s global font, and applies the style to all cells that already contain data. | Save the modified workbook to a new file after globally setting a sans‑serif font using Aspose.Cells style flags.
// Common Searches: Aspose.Cells C# change default workbook font to Arial and apply to existing cells | how to set a global sans‑serif font for all worksheets in an Excel file using Aspose.Cells | apply default style font to used range of each sheet with Aspose.Cells .NET | programmatically update Excel theme primary font in C# Aspose.Cells | save workbook after changing default font across all cells Aspose.Cells
// Tags: default font update Aspose.Cells C# | apply global font style used range | set workbook default style Arial | propagate default style to existing cells | save workbook with modified theme font

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx file, sets the workbook's DefaultStyle.Font to a sans‑serif type (Arial), iterates through each worksheet to apply this default style to the entire used range using a StyleFlag that targets the font, ensures the output directory exists, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the default (global) font to a sans‑serif font (e.g., Arial)
            workbook.DefaultStyle.Font.Name = "Arial";

            // Optionally, apply the default style to all existing cells
            // This ensures the change affects already formatted cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Apply the default style to the entire used range of the sheet
                var usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange != null)
                {
                    Style defaultStyle = workbook.DefaultStyle;
                    usedRange.ApplyStyle(defaultStyle, new StyleFlag { Font = true });
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated font applied globally
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
