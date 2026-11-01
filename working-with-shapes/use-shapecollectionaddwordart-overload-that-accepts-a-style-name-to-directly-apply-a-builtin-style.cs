// Title: Add WordArt with a preset style to an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, uses ShapeCollection.AddWordArt to insert WordArt with a built‑in style, and saves the workbook as an .xlsx file. | Demonstrate calling the AddWordArt method with a style enumeration to set a preset WordArt appearance in Aspose.Cells. | Provide a concise example that sets the position, size, and style of WordArt on a worksheet using the Aspose.Cells C# API.
// Common Searches: asp.net cells c# add wordart preset style worksheet example | shapecollection addwordart overload style parameter Aspose.Cells | how to use WordArtStyle1 with Aspose.Cells in C# | create styled wordart in excel file using Aspose.Cells .NET | apply built‑in wordart formatting to .xlsx via Aspose.Cells C#
// Tags: apply built‑in WordArt style Aspose.Cells | C# insert WordArt into Excel workbook | Aspose.Cells shape collection WordArt example | Excel .xlsx styled WordArt via .NET | preset WordArt style usage in C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample creates a new workbook, accesses the first worksheet, and uses ShapeCollection.AddWordArt with PresetWordArtStyle.WordArtStyle1 to insert styled WordArt text, then saves the file as WordArtWithStyle.xlsx.
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

                // Add WordArt using a preset style
                // Parameters: style, text, upper left row, upper left column,
                // lower right row, lower right column, height, width
                Shape wordArt = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "Aspose.Cells WordArt",
                    2,   // Upper left row (zero‑based)
                    1,   // Upper left column (zero‑based)
                    6,   // Lower right row
                    5,   // Lower right column
                    100, // Height (points)
                    300  // Width (points)
                );

                // Save the workbook to a file
                string outputPath = "WordArtWithStyle.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
