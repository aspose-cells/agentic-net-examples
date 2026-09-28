// Title: Add a semi‑transparent diagonal WordArt watermark to every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through all worksheets in a workbook and inserts a TextEffect shape with custom text as a semi‑transparent diagonal watermark using Aspose.Cells. | Show how to compute worksheet dimensions, center a WordArt shape, set its rotation to -45°, and apply 50% transparency with the Aspose.Cells API. | Provide a complete example that loads an existing .xlsx file, adds the watermark to each sheet, and saves the modified workbook to a new file in C#.
// Common Searches: Aspose.Cells C# add diagonal WordArt watermark to each sheet in an existing Excel file | how to set transparency on a TextEffect shape using Aspose.Cells .NET | center and rotate WordArt shape programmatically with Aspose.Cells workbook | iterate through worksheets and apply same watermark using Aspose.Cells API
// Tags: Aspose.Cells add WordArt watermark to worksheets | C# semi-transparent Excel watermark using TextEffect | Aspose.Cells rotate shape diagonal | center shape on worksheet dimensions Aspose.Cells | iterate workbook worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing .xlsx workbook, iterates over every worksheet, adds a TextEffect (WordArt) shape containing the text "CONFIDENTIAL", centers it, rotates it -45° for a diagonal layout, sets 50% transparency, removes the outline, and saves the watermarked workbook as a new file.
class WatermarkExample
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a WordArt (TextEffect) shape to the worksheet
                // Provide height and width parameters required by the overload
                Shape wordArt = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // Preset effect
                    "CONFIDENTIAL",                 // Watermark text
                    "Arial",                        // Font name
                    72,                             // Font size
                    true,                           // Bold
                    false,                          // Italic
                    0,                              // Upper left row
                    0,                              // Upper left column
                    0,                              // Top offset (in points)
                    0,                              // Left offset (in points)
                    200,                            // Height (in points)
                    500);                           // Width (in points)

                // Set the size of the WordArt shape (width and height in points)
                wordArt.Width = 500;
                wordArt.Height = 200;

                // Approximate sheet dimensions in points
                int sheetWidth = sheet.Cells.MaxColumn * 64;
                int sheetHeight = sheet.Cells.MaxRow * 20;

                // Center the shape on the sheet
                wordArt.Left = (sheetWidth - wordArt.Width) / 2;
                wordArt.Top = (sheetHeight - wordArt.Height) / 2;

                // Rotate the watermark for a diagonal appearance
                wordArt.RotationAngle = -45;

                // Make the WordArt semi‑transparent (0 = opaque, 1 = fully transparent)
                wordArt.Fill.Transparency = 0.5; // 50% transparent

                // Optionally remove the outline by setting line weight to zero
                wordArt.Line.Weight = 0;
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Watermarked workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
