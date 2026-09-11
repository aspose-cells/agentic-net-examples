// Title: Apply Bold Wave WordArt formatting to TextEffect shapes on worksheets prefixed with "Report" using Aspose.Cells for .NET
// AI Prompts: Iterate over all worksheets whose names begin with "Report", locate each TextEffect shape, set FontBold to true and FontSize to 36, then save the workbook. | Filter worksheets by a name prefix, find every WordArt (TextEffect) shape, apply a bold 36‑point style, and export the updated Excel file.
// Common Searches: Aspose.Cells C# apply bold WordArt to shapes on worksheets that start with Report | How to set FontBold and FontSize for TextEffect shapes in specific Excel sheets using Aspose.Cells | C# code to format WordArt shapes only on worksheets with a name prefix | Conditional shape formatting by worksheet name in Aspose.Cells for .NET
// Tags: bold wordart texteffect shapes aspnet | filter worksheets by name prefix aspose.cells | set texteffect font properties c# | conditional shape formatting excel workbook | wordart formatting aspose cells c#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads input.xlsx, processes only worksheets whose names start with "Report", sets each TextEffect shape's WordArt font to bold and size 36, and saves the modified workbook to output.xlsx.
class ApplyWordArtStyle
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process only worksheets whose names start with "Report"
                if (sheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Apply formatting only to shapes that are TextEffect (WordArt)
                        if (shape.TextEffect != null)
                        {
                            // Apply typical WordArt-like formatting manually
                            shape.TextEffect.FontBold = true;
                            shape.TextEffect.FontSize = 36;
                            // FontColor property may not be available in some versions; omitted for compatibility
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display the error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
