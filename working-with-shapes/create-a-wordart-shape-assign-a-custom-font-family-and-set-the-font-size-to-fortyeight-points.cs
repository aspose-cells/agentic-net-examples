// Title: Add WordArt with a custom font and 48‑pt size using Aspose.Cells for .NET (C#)
// Description: This example shows how to create a new Workbook, add a WordArt shape via ShapeCollection.AddTextEffect, assign the custom font "MyCustomFont" at 48 points, adjust the TextEffectFormat, and save the file as WordArtCustomFont48pt.xlsx.
// Keywords: Aspose.Cells C# WordArt | Add WordArt shape | custom font WordArt Aspose.Cells | 48 point font size | TextEffectFormat API | Excel shape collection .NET | Aspose.Cells example GitHub | C# Excel WordArt code
// Common Searches: How to create WordArt in Excel with Aspose.Cells C# | Aspose.Cells set WordArt font size to 48 points | WordArt custom font example Aspose.Cells .NET | AddTextEffect usage C# | Change WordArt text effect programmatically
// Developer Intent: Insert a WordArt shape into an Excel worksheet using a specific custom font family and a 48‑point font size.
// Use Cases: Generate branded report titles with the company’s custom font for consistent visual identity. | Add decorative headings to spreadsheet templates where a large, eye‑catching font is required. | Automate placement of WordArt across multiple worksheets while preserving exact font styling.
// AI Prompts: Provide C# code to change the fill and outline colors of a WordArt shape created with Aspose.Cells. | Show how to align a WordArt shape to the center of a given cell range in an Excel worksheet. | Explain how to list all MsoPresetTextEffect values and select one based on user input in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtDemo
{
    // This example shows how to create a new Workbook, add a WordArt shape via ShapeCollection.AddTextEffect, assign the custom font "MyCustomFont" at 48 points, adjust the TextEffectFormat, and save the file as WordArtCustomFont48pt.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add a WordArt shape using AddTextEffect.
            // Parameters:
            //   effect: preset text effect (choose any, e.g., TextEffect1)
            //   text: the WordArt text
            //   fontName: custom font family
            //   size: initial font size (set to 48)
            //   fontBold, fontItalic: false
            //   topRow, top, leftColumn, left: position in cells/pixels
            //   height, width: size of the shape in pixels
            Shape wordArt = shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Custom WordArt",
                "MyCustomFont",   // custom font family
                48,               // font size in points
                false,
                false,
                2,    // top row index
                0,    // vertical offset (pixels)
                2,    // left column index
                0,    // horizontal offset (pixels)
                200,  // height (pixels)
                400   // width (pixels)
            );

            // If further adjustments are needed, modify the TextEffect properties
            TextEffectFormat textEffect = wordArt.TextEffect;
            textEffect.FontName = "MyCustomFont";
            textEffect.FontSize = 48;

            // Save the workbook with the WordArt shape
            workbook.Save("WordArtCustomFont48pt.xlsx");
        }
    }
}
