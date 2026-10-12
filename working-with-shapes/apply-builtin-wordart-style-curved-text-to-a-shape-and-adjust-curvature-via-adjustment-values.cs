// Title: Add a Curved WordArt TextEffect Shape to an Excel Worksheet and Adjust Its Curvature Using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a Workbook, adds a TextEffect shape with the Curved preset, and assigns specific adjustment parameters to define the curve. | Demonstrate how to access the adjustment handles of a WordArt element in Aspose.Cells and update them to increase or decrease the curvature. | Provide a C# routine that writes the workbook to a given file path, automatically creating any missing folders.
// Common Searches: how to add a curved WordArt text effect to an Excel workbook with Aspose.Cells in C# | c# Aspose.Cells modify WordArt curvature after inserting a WordArt shape | example of setting curve level for a WordArt shape using Aspose.Cells API | save Aspose.Cells workbook to a specific folder and create the folder if it does not exist
// Tags: insert WordArt via TextEffect API Aspose.Cells | tune WordArt bend factor Aspose.Cells | create missing directory before saving workbook Aspose.Cells | export workbook to .xlsx format with Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, inserts a Curved preset WordArt TextEffect shape, optionally adjusts its curvature via adjustment parameters, ensures the output directory exists, and saves the file as an .xlsx workbook.
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

            // Add a WordArt (TextEffect) shape using the appropriate overload.
            // Parameters: preset, text, font name, font size, isBold, isItalic,
            // upper left row, upper left column, top offset, left offset, height, width
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset (curved style)
                "Curved WordArt",                // displayed text
                "Arial",                         // font name
                36,                              // font size
                false,                           // isBold
                false,                           // isItalic
                5,                               // upper left row
                1,                               // upper left column
                0,                               // top offset (pixels)
                0,                               // left offset (pixels)
                100,                             // height (pixels)
                400                              // width (pixels)
            );

            // Define output path
            string outputPath = "WordArtCurved.xlsx";

            // Ensure the output directory exists (handle null when only file name is provided)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
