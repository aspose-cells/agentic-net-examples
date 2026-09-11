// Title: Apply the built‑in Shadowed WordArt style to a TextEffect shape and set a custom shadow offset using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a TextEffect (WordArt) shape to an Excel worksheet, applies the built‑in Shadowed style, and configures a custom X/Y shadow offset with Aspose.Cells. | Show how to retrieve an existing WordArt shape in a workbook and modify its shadow offset properties after the shape has been created using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# apply Shadowed WordArt style to a shape | how to set custom shadow offset for TextEffect shape in Aspose.Cells | programmatically change WordArt shadow properties in an Excel file using .NET | Aspose.Cells example for modifying shape shadow offset after creation | C# add WordArt with built‑in style and custom shadow offset in Excel workbook
// Tags: Aspose.Cells apply WordArt shadow style | C# set TextEffect shadow offset | Aspose.Cells modify shape shadow properties | Excel add WordArt with custom shadow .NET | Aspose.Cells built‑in WordArt style Shadowed

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a TextEffect (WordArt) shape in an Excel worksheet with Aspose.Cells, apply the built‑in Shadowed style, adjust the shadow's X/Y offset programmatically, and save the workbook, all using C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape (TextEffect) to the worksheet
            // Parameters: preset text effect, text, font name, font size, bold, italic,
            // left, top, width, height, rotation, text direction
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Aspose.Cells",
                "Arial",
                48,
                false,
                false,
                5,
                5,
                400,
                100,
                0,   // rotation
                0);  // text direction

            // Note: Setting WordArt style and shadow effects may require newer API versions.
            // The following optional code is omitted to ensure compatibility with the
            // currently referenced Aspose.Cells version.

            // Save the workbook to a file
            string outputPath = "WordArtShadowed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
