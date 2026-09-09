// Title: Insert WordArt with a built‑in preset style into an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new workbook, adds a WordArt shape using the PresetWordArtStyle.WordArtStyle1 preset, sets the font size to 24, and saves the file as WordArtStyleExample.xlsx with Aspose.Cells. | Write a C# snippet that uses the ShapeCollection.AddWordArt overload to place WordArt spanning rows 2‑6 and columns 1‑5, specifies a width of 400 points and a height of 100 points, applies a built‑in style, and outputs the workbook.
// Common Searches: Aspose.Cells C# add WordArt with preset style to specific cell range | How to use ShapeCollection.AddWordArt overload for built‑in WordArt styles in .NET | C# example of applying WordArtStyle1 to a shape in an Excel file with Aspose.Cells | Set font size of WordArt shape created by Aspose.Cells | Save workbook containing WordArt shape using Aspose.Cells for .NET
// Tags: ShapeCollection.AddWordArt preset style | apply WordArtStyle1 Aspose.Cells | WordArt shape font size C# | save workbook with WordArt Aspose.Cells | Excel worksheet WordArt built‑in style .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtExample
{
    // Demonstrates creating a new workbook, accessing the first worksheet, adding a WordArt shape with the WordArtStyle1 preset via ShapeCollection.AddWordArt, adjusting its font size, and saving the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Get the shape collection of the worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Add WordArt with a built‑in style.
                // Parameters: style, text, upper left row, upper left column,
                // lower right row, lower right column, width, height
                Shape wordArt = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "Aspose.Cells WordArt",
                    2, // upper left row (zero‑based)
                    1, // upper left column (zero‑based)
                    6, // lower right row
                    5, // lower right column
                    400, // width in points
                    100  // height in points
                );

                // Optionally adjust the font size
                wordArt.Font.Size = 24;

                // Save the workbook to a file
                string outputPath = "WordArtStyleExample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
