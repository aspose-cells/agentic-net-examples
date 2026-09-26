// Title: Create a second WordArt shape with Simple Fill style and set a custom font size in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a WordArt shape using the SimpleFill preset, sets its caption to "Second WordArt", and applies a 36‑point text size. | Write a function that inserts an additional WordArt object into a worksheet, selects the Simple Fill style, and changes the text size programmatically.
// Common Searches: Aspose.Cells C# add WordArt with SimpleFill preset and adjust text size | how to create multiple WordArt objects in an Excel file using Aspose.Cells .NET | set text size for WordArt shape in Aspose.Cells example code | C# example for applying Simple Fill style to WordArt in Excel workbook
// Tags: Aspose.Cells add WordArt shape C# | SimpleFill WordArt preset Aspose.Cells | custom text size WordArt Aspose.Cells | multiple WordArt objects Excel .NET | WordArt styling with preset Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a second WordArt shape with the Simple Fill preset, sets its text to "Second WordArt" and applies a 36‑point font size, then saves the file as WordArtExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape to the worksheet.
            // Use a cast to the enum value (0 = SimpleFill) to avoid version‑specific enum member issues.
            Shape wordArtShape = worksheet.Shapes.AddWordArt(
                (PresetWordArtStyle)0,          // SimpleFill style
                "Second WordArt",
                5, 0, 10, 5,
                200, 100);

            // Set a custom font size for the WordArt text
            wordArtShape.TextEffect.FontSize = 36;

            // Determine output path and ensure the directory exists
            string outputPath = "WordArtExample.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
