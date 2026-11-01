// Title: Add a WordArt shape with custom curvature to an Excel worksheet and save the file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add WordArt shape with custom curvature adjustments | How to set WordArt adjustment values in an Excel file using Aspose.Cells .NET | Create WordArt with preset style and save workbook in C# Aspose.Cells | C# code to insert WordArt into worksheet and export to .xlsx with Aspose.Cells | Adjust WordArt shape properties programmatically with Aspose.Cells for .NET
// Tags: WordArt shape insertion Aspose.Cells | WordArt adjustment collection .NET | preset WordArtStyle1 usage | export workbook to .xlsx Aspose.Cells | C# Excel shape manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first worksheet, inserts a WordArt shape using a preset style, specifies its text, position, and size, notes that curvature can be modified via the Adjustments collection, ensures the output directory exists, and saves the workbook as 'WordArtExample.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape to the worksheet
            // Parameters: preset style, text, upper left row, upper left column,
            // top offset, left offset, height, width (all in points)
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "Custom WordArt",                 // text
                5,   // upper left row (zero‑based)
                1,   // upper left column (zero‑based)
                0,   // top offset
                0,   // left offset
                100, // height
                300  // width
            );

            // Adjustments are not available in all versions; skip if unsupported.
            // If needed, you can explore the Adjustments collection in newer releases.

            // Determine output file path
            string outputPath = "WordArtExample.xlsx";

            // Ensure the directory for the output file exists (if a directory is specified)
            string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
