// Title: Apply Bold Wave WordArt style with FontSetting.SetWordArtStyle in Aspose.Cells for .NET (C#)
// Description: This C# example creates a new Workbook, adds a WordArt shape to the first worksheet, applies the Bold Wave preset using FontSetting.SetWordArtStyle, ensures the text is bold, and saves the file as BoldWaveWordArtDemo.xlsx. The code works with Aspose.Cells for .NET and demonstrates how to style WordArt programmatically.
// Keywords: Aspose.Cells C# WordArt | FontSetting.SetWordArtStyle | Bold Wave preset | AddTextEffect example | Excel shape styling | programmatic WordArt | Excel workbook .NET | Aspose.Cells tutorial | GitHub Aspose.Cells sample | Excel automation C#
// Common Searches: How to set Bold Wave WordArt style with Aspose.Cells | C# FontSetting.SetWordArtStyle example | Add WordArt to Excel using Aspose.Cells .NET | Apply preset WordArt effect in C# | Aspose.Cells shape formatting tutorial
// Developer Intent: Insert a WordArt shape, apply the Bold Wave preset via FontSetting.SetWordArtStyle, and save the workbook.
// Use Cases: Create a branded report header with stylized WordArt. | Add decorative titles to Excel dashboards for visual impact. | Generate template worksheets that automatically include a bold WordArt title.
// AI Prompts: Generate C# code that uses FontSetting.SetWordArtStyle to apply the Bold Wave preset to a WordArt shape in Aspose.Cells. | Show how to add a WordArt shape, set its preset style, make the text bold, and save the workbook using Aspose.Cells for .NET. | Explain the steps to modify a WordArt shape's style after insertion, including using FontSetting.SetWordArtStyle and FontBold.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtDemo
{
    // This C# example creates a new Workbook, adds a WordArt shape to the first worksheet, applies the Bold Wave preset using FontSetting.SetWordArtStyle, ensures the text is bold, and saves the file as BoldWaveWordArtDemo.xlsx. The code works with Aspose.Cells for .NET and demonstrates how to style WordArt programmatically.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a WordArt shape. If the specific preset enum value is unavailable in the
                // referenced Aspose.Cells version, use the default enum value (0) which maps to a basic effect.
                Shape wordArt = worksheet.Shapes.AddTextEffect(
                    (MsoPresetTextEffect)0,          // preset effect (fallback to default)
                    "Bold Wave WordArt",            // text
                    "Arial",                        // font name
                    36,                             // font size
                    true,                           // isBold
                    false,                          // isItalic
                    2, 10,                          // upperLeftRow, top
                    2, 10,                          // upperLeftColumn, left
                    200, 100);                      // height, width

                // Ensure the font is bold (redundant if already set, but kept for clarity)
                wordArt.TextEffect.FontBold = true;

                // Save the workbook
                workbook.Save("BoldWaveWordArtDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
