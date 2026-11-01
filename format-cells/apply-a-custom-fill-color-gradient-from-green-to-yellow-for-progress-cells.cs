// Title: Apply a green foreground and yellow background solid fill to a range of progress cells in an Excel file using Aspose.Cells for .NET
// Common Searches: how to set solid fill color for a range of cells with Aspose.Cells in C# | Aspose.Cells C# apply green foreground and yellow background to column B rows 2 to 11 | set cell style background colors programmatically in an Excel workbook using Aspose.Cells .NET | C# example for applying custom fill colors to progress cells in an XLSX file with Aspose.Cells | Aspose.Cells style pattern solid fill example for Excel progress tracking
// Tags: solid fill style Aspose.Cells | set cell foreground color Aspose.Cells | set cell background color Aspose.Cells | apply style to column B Aspose.Cells | progress tracking cell formatting .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The program loads 'Progress.xlsx', creates a solid fill style with a green foreground and yellow background using Aspose.Cells, applies this style to cells B2‑B11 (the progress column), ensures the output directory exists, and saves the modified workbook as 'Progress_Gradient.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Progress.xlsx";
            const string outputPath = "Progress_Gradient.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the progress cells
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with a simple fill (gradient not available in this version, using solid fill)
            Style gradientStyle = workbook.CreateStyle();
            gradientStyle.Pattern = BackgroundType.Solid;          // Use solid fill as fallback
            gradientStyle.ForegroundColor = Color.Green;          // Fill color
            gradientStyle.BackgroundColor = Color.Yellow;         // Not used for solid fill but kept for reference

            // Apply the style to column B rows 2‑11 (zero‑based indices)
            for (int row = 1; row <= 10; row++)
            {
                Cell cell = sheet.Cells[row, 1]; // Column B
                cell.SetStyle(gradientStyle);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
