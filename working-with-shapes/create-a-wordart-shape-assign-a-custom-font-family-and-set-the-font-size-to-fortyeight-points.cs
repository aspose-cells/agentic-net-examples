// Title: Add a WordArt shape with Comic Sans MS font and 48‑point size to an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a WordArt object into the first worksheet, sets its font to 'Comic Sans MS', and applies a 48‑point size using Aspose.Cells. | Show how to configure the Font.Name and Font.Size properties of a WordArt shape after adding it with AddWordArt in Aspose.Cells. | Demonstrate saving an Excel file after customizing a WordArt shape's font family and point size with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to set WordArt font name and size in Excel workbook | example code for adding WordArt with custom font using Aspose.Cells .NET | change WordArt font to Comic Sans MS and set 48 point size with Aspose.Cells
// Tags: add-wordart shape aspocells | customize-wordart font family c# | set-wordart font size 48 points | aspocells shape font properties | save-excel workbook aspocells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// // Creates a new workbook, inserts a WordArt shape with Comic Sans MS font at 48‑point size, and saves it as WordArtExample.xlsx using Aspose.Cells.
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

            // Add a WordArt shape with sample text
            // Parameters: style, text, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Sample WordArt",
                1,    // upper left row (zero‑based)
                0,    // upper left column
                5,    // top offset (pixels)
                5,    // left offset (pixels)
                200,  // height (pixels)
                400); // width (pixels)

            // Assign a custom font family to the WordArt
            wordArt.Font.Name = "Comic Sans MS";

            // Set the font size to 48 points
            wordArt.Font.Size = 48;

            // Save the workbook to a file
            string outputPath = "WordArtExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
