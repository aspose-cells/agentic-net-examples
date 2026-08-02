// Title: Group and lock WordArt shapes in an Excel workbook with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add two WordArt (text‑effect) shapes to a worksheet, combine them into a GroupShape using ShapeCollection.Group, set the group's IsLocked flag to protect its layout, and save the result as GroupedWordArt.xlsx.
// Keywords: Aspose.Cells C# group shapes | lock shape group Aspose.Cells | WordArt Excel .NET | ShapeCollection.Group example | IsLocked property Excel | protect Excel shapes programmatically | Aspose.Cells shape grouping tutorial | C# add WordArt to workbook
// Common Searches: Aspose.Cells group WordArt shapes | How to lock a shape group in Excel using C# | ShapeCollection.Group method Aspose.Cells | Prevent moving grouped shapes in Aspose.Cells | C# example for WordArt grouping and protection
// Developer Intent: Create a grouped WordArt object and lock it to keep the layout fixed in an Excel file.
// Use Cases: Design a fixed banner made of multiple WordArt elements that stays together when users edit the sheet. | Secure a composite logo built from several shapes so it cannot be moved or resized accidentally. | Maintain consistent positioning of decorative text across printed reports by grouping and locking the shapes.
// AI Prompts: Generate C# code to ungroup a locked WordArt GroupShape with Aspose.Cells. | Show how to add an additional WordArt shape to an existing locked group without releasing the lock. | Provide an example of applying protection settings to individual shapes inside a group (e.g., IsLocked, IsPrintable).

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtGroupDemo
{
    // Demonstrates how to add two WordArt (text‑effect) shapes to a worksheet, combine them into a GroupShape using ShapeCollection.Group, set the group's IsLocked flag to protect its layout, and save the result as GroupedWordArt.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Get the shapes collection of the worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Add two WordArt (text effect) shapes
                // Parameters: preset effect, text, font name, font size, bold, italic,
                // upper left row, upper left column, lower right row, lower right column, height, width
                Shape wordArt1 = shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Hello",
                    "Arial",
                    36,
                    false,
                    false,
                    2,   // upper left row
                    2,   // upper left column
                    4,   // lower right row
                    4,   // lower right column
                    100, // height (pixels)
                    200  // width (pixels)
                );

                Shape wordArt2 = shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect2,
                    "World",
                    "Arial",
                    36,
                    false,
                    false,
                    5,   // upper left row
                    5,   // upper left column
                    7,   // lower right row
                    7,   // lower right column
                    100, // height (pixels)
                    200  // width (pixels)
                );

                // Group the two WordArt shapes
                GroupShape group = shapes.Group(new Shape[] { wordArt1, wordArt2 });

                // Lock the group to maintain layout integrity
                group.IsLocked = true;

                // Save the workbook
                string outputPath = "GroupedWordArt.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
