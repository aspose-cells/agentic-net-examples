// Title: How to apply the Bold Wave preset WordArt style to a rectangle shape using Aspose.Cells FontSetting.SetWordArtStyle in C#
// Common Searches: C# Aspose.Cells set Bold Wave WordArt style on a shape | Apply preset WordArt style to Excel shape with FontSetting.SetWordArtStyle | Aspose.Cells example for adding WordArt style to rectangle shape in workbook | Using PresetWordArtStyle.WordArtStyle7 with Aspose.Cells C# | How to style shape text with Bold Wave WordArt in Aspose.Cells
// Tags: Bold Wave WordArt style with Aspose.Cells | FontSetting.SetWordArtStyle method C# | Excel rectangle shape text WordArt | PresetWordArtStyle.WordArtStyle7 implementation | Aspose.Cells shape text formatting example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace ApplyWordArtStyleDemo
{
    // The sample creates a new workbook, adds a rectangle shape, sets its TextBody text, applies the Bold Wave preset WordArt style using FontSetting.SetWordArtStyle (PresetWordArtStyle.WordArtStyle7), and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting workbook will be saved
            const string outputPath = "BoldWaveWordArtDemo.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Add a shape that can contain text (a rectangle)
                // Parameters: drawing type, upper left row, top offset (pixels),
                // upper left column, left offset (pixels), height (pixels), width (pixels)
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    2,   // upper left row index
                    10,  // top offset in pixels
                    2,   // upper left column index
                    10,  // left offset in pixels
                    100, // height in pixels
                    300  // width in pixels
                );

                // 3. Access the FontSettingCollection (TextBody) of the shape
                FontSettingCollection textBody = shape.TextBody;

                // 4. Set the text that will receive the WordArt style
                textBody.Text = "Bold Wave WordArt Example";

                // 5. Apply the preset WordArt style.
                // In Excel the style named "Bold Wave" corresponds to PresetWordArtStyle.WordArtStyle7.
                // If a newer version of Aspose.Cells adds a dedicated enum value, replace it accordingly.
                textBody.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);

                // 6. Save the workbook to see the effect in Excel
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Open the file in Excel to view the shape with the 'Bold Wave' WordArt style applied.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while applying the WordArt style:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
