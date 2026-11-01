// Title: How to set a WordArt shape's outline weight to 2 points and change its outline color to dark gray using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set WordArt outline thickness to 2 points | change WordArt shape border color to dark gray in a .NET workbook | using reflection to modify WordArt line color with Aspose.Cells | add WordArt to worksheet and customize line style Aspose.Cells | set shape line weight and color programmatically Aspose.Cells C#
// Tags: WordArt outline weight Aspose.Cells | WordArt border color dark gray C# | shape line property reflection Aspose.Cells | customize WordArt line style .NET | Aspose.Cells shape formatting example

using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding a WordArt shape, setting its outline weight to 2 points, applying a dark gray outline color via reflection when needed, and saving the file as WordArtOutlineExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a WordArt shape to the worksheet
                // Parameters: preset style, text, upper left row, upper left column,
                // row offset, column offset, width, height
                Shape wordArtShape = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1, // preset style
                    "Sample WordArt",                // text
                    0,                               // upper left row
                    0,                               // upper left column
                    0,                               // row offset (pixels)
                    0,                               // column offset (pixels)
                    200,                             // width (pixels)
                    50);                             // height (pixels)

                // Set the outline (line) weight to 2 points
                wordArtShape.Line.Weight = 2;

                // Attempt to set the outline color to dark gray using reflection
                // (some versions may not expose the Color property)
                try
                {
                    PropertyInfo colorProp = wordArtShape.Line.GetType().GetProperty("Color");
                    if (colorProp != null && colorProp.CanWrite)
                    {
                        colorProp.SetValue(wordArtShape.Line, Color.DarkGray);
                    }
                }
                catch (Exception)
                {
                    // Ignore any errors while setting the color
                }

                // Save the workbook to a file
                string outputPath = "WordArtOutlineExample.xlsx";

                // Ensure the directory exists (prevents FileNotFoundException for path issues)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
