// Title: C# – Change WordArt Font Color to Dark Blue Across All Worksheets with Aspose.Cells
// Description: Loads an Excel workbook, iterates each worksheet and its Shapes collection, identifies WordArt objects via IsWordArt, sets their Font.Color to Color.DarkBlue, and saves the updated file.
// Keywords: Aspose.Cells | C# | WordArt | font color | dark blue | IsWordArt | iterate shapes | Excel automation | bulk style update | shape.Font.Color
// Common Searches: Aspose.Cells change WordArt color .NET | set WordArt text color in Excel using C# | loop through shapes and modify WordArt font | bulk update WordArt font color Aspose | C# code to change WordArt to dark blue
// Developer Intent: Apply a dark‑blue font color to every WordArt shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enforce corporate branding by converting all WordArt captions to the official dark‑blue shade in multi‑sheet reports. | Prepare a template workbook where WordArt text must match a predefined color before distribution to end users. | Automate the correction of imported WordArt objects that appear in default colors, ensuring visual consistency across all worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that changes WordArt font color to a custom RGB value instead of DarkBlue. | Explain how to filter shapes by type (WordArt, picture, chart) and modify their properties using Aspose.Cells for .NET. | Provide a snippet that logs the name or index of each WordArt shape whose color was updated during processing.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtColorChange
{
    // Loads an Excel workbook, iterates each worksheet and its Shapes collection, identifies WordArt objects via IsWordArt, sets their Font.Color to Color.DarkBlue, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Check if the shape is a WordArt object
                    if (shape.IsWordArt)
                    {
                        // Change the text color to dark blue
                        shape.Font.Color = Color.DarkBlue;
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
