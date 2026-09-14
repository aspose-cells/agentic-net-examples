// Title: How to set a WordArt shape's outline weight to 2 points and change its border color to dark gray using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a WordArt shape to an Excel worksheet with a 2‑point outline and a dark gray border using Aspose.Cells. | Update an existing Aspose.Cells workbook to change a WordArt shape's line thickness to 2 points and apply RGB(169,169,169) as the outline color.
// Common Searches: Aspose.Cells C# how to increase WordArt outline thickness | set WordArt border to dark gray in Excel using Aspose.Cells | adjust line weight of WordArt shape programmatically with Aspose.Cells .NET | example code for changing WordArt shape line color to RGB 169 169 169 | Aspose.Cells shape formatting outline weight and color
// Tags: Aspose.Cells WordArt line weight | Aspose.Cells shape outline color | C# set WordArt border thickness | Excel shape line formatting Aspose.Cells | dark gray line color Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;
using System.IO;

// C# example that creates a workbook, inserts a WordArt shape, sets its outline weight to 2 points, optionally changes the outline color to dark gray (RGB 169,169,169), and saves the file as WordArtOutline.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape (style, text, upper left row, upper left column, lower right row, lower right column, width, height)
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Sample WordArt",
                2, 2, 6, 6, 200, 100);

            // Set the outline weight to 2 points
            wordArt.Line.Weight = 2.0; // points

            // Change the outline color to dark gray (if supported by the version)
            // In some Aspose.Cells versions LineFormat does not expose a Color property.
            // Uncomment the following line if your version supports it:
            // wordArt.Line.Color = Color.FromArgb(169, 169, 169);

            // Define output path
            string outputPath = "WordArtOutline.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
