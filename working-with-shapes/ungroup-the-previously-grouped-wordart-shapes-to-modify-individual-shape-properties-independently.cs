// Title: How to ungroup WordArt shapes and edit each shape’s fill color, rotation, and text with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# ungroup WordArt shapes and change fill color | modify individual WordArt shape rotation after grouping using Aspose.Cells | C# example for updating WordArt text after ungrouping in Aspose.Cells | Ungroup grouped WordArt in Aspose.Cells and set shape properties programmatically | iterate WordArt shapes in Aspose.Cells to apply fill and rotation changes
// Tags: ungroup WordArt shapes Aspose.Cells | set WordArt fill color C# | rotate WordArt shape Aspose.Cells | modify WordArt text programmatically | group and ungroup shapes API Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace UngroupWordArtDemo
{
    // The sample creates a workbook, adds two WordArt shapes, groups them, then uses the Ungroup method to separate the shapes. It iterates over each WordArt shape to change its fill color, rotation angle, and text before saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting workbook will be saved
            string outputPath = "UngroupWordArtDemo.xlsx";

            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Add two WordArt shapes to the worksheet
                // -----------------------------------------------------------------
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape wordArt1 = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "First WordArt",
                    2,   // topRow
                    20,  // top (pixels)
                    2,   // leftColumn
                    20,  // left (pixels)
                    100, // height (pixels)
                    300  // width (pixels)
                );

                Shape wordArt2 = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle2,
                    "Second WordArt",
                    5,   // topRow
                    150, // top (pixels)
                    5,   // leftColumn
                    150, // left (pixels)
                    100, // height (pixels)
                    300  // width (pixels)
                );

                // Verify that the shapes are indeed WordArt
                Console.WriteLine($"WordArt1.IsWordArt = {wordArt1.IsWordArt}");
                Console.WriteLine($"WordArt2.IsWordArt = {wordArt2.IsWordArt}");

                // -----------------------------------------------------------------
                // 3. Group the two WordArt shapes
                // -----------------------------------------------------------------
                Shape[] shapesToGroup = new Shape[] { wordArt1, wordArt2 };
                GroupShape group = worksheet.Shapes.Group(shapesToGroup);

                Console.WriteLine($"Group created. IsGroup = {group.IsGroup}");

                // -----------------------------------------------------------------
                // 4. Ungroup the previously grouped WordArt shapes
                // -----------------------------------------------------------------
                // Ungroup returns void; after this call the individual shapes are back in the collection
                group.Ungroup();

                Console.WriteLine("Shapes have been ungrouped.");

                // -----------------------------------------------------------------
                // 5. Modify individual shape properties independently
                // -----------------------------------------------------------------
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Ensure we are working with a WordArt shape
                    if (shape.IsWordArt)
                    {
                        // Example modifications:
                        // - Change fill color
                        // - Rotate the shape
                        // - Change the text
                        shape.Fill.SolidFill.Color = Color.LightBlue;
                        shape.RotationAngle = 15; // rotate 15 degrees
                        shape.Text = shape.Text + " (modified)";
                    }
                }

                // -----------------------------------------------------------------
                // 6. Save the workbook
                // -----------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
