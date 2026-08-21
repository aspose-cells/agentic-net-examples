// Title: Ungroup WordArt Shapes and Edit TextEffect Individually with Aspose.Cells for .NET (C#)
// Description: This Aspose.Cells for .NET example creates a workbook, adds two WordArt shapes, groups them, then uses GroupShape.Ungroup to separate the shapes. After ungrouping each WordArt's TextEffect (font, size, bold, italic) is modified independently before saving the file.
// Keywords: Aspose.Cells | C# | WordArt | Ungroup shapes | GroupShape.Ungroup | TextEffectFormat | shape formatting | Excel API | code example | GitHub
// Common Searches: Aspose.Cells ungroup WordArt C# | how to edit TextEffect after grouping shapes | GroupShape.Ungroup example Aspose.Cells | modify individual WordArt font properties .NET | Aspose.Cells shape editing tutorial
// Developer Intent: The developer needs to separate previously grouped WordArt objects so each shape’s TextEffect (font name, size, bold, italic) can be changed independently.
// Use Cases: Design a spreadsheet layout with grouped WordArt headings, then apply distinct font styles to each heading. | Generate a report where WordArt titles are positioned together for alignment, but require individual styling before export. | Programmatically adjust bold, italic, and font size of separate WordArt shapes after they have been grouped for layout purposes.
// AI Prompts: Show C# code to ungroup a GroupShape in Aspose.Cells and change the TextEffect of each WordArt shape. | Provide an Aspose.Cells example that groups multiple WordArt objects, then ungroups them to set different font attributes. | Explain how to modify font style, size, and bold/italic settings of individual WordArt after using GroupShape.Ungroup.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtUngroupDemo
{
    // This Aspose.Cells for .NET example creates a workbook, adds two WordArt shapes, groups them, then uses GroupShape.Ungroup to separate the shapes. After ungrouping each WordArt's TextEffect (font, size, bold, italic) is modified independently before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two WordArt shapes to the worksheet
                Shape wordArt1 = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1, // preset style
                    "First WordArt",                  // text
                    2, 0,                             // upper left row, column
                    100, 300,                         // height, width (in pixels)
                    0, 0);                            // row offset, column offset

                Shape wordArt2 = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle2,
                    "Second WordArt",
                    6, 0,
                    100, 300,
                    0, 0);

                // Group the two WordArt shapes
                GroupShape group = worksheet.Shapes.Group(new Shape[] { wordArt1, wordArt2 });

                // Ungroup the shapes so they can be edited individually
                group.Ungroup();

                // After ungrouping, modify the TextEffect of each WordArt shape independently
                if (wordArt1.IsWordArt)
                {
                    TextEffectFormat effect1 = wordArt1.TextEffect;
                    effect1.FontBold = true;
                    effect1.FontItalic = true;
                    effect1.FontName = "Arial";
                    effect1.FontSize = 16;
                    // Underline not supported directly; can be handled via other formatting if needed
                }

                if (wordArt2.IsWordArt)
                {
                    TextEffectFormat effect2 = wordArt2.TextEffect;
                    effect2.FontBold = false;
                    effect2.FontItalic = false;
                    effect2.FontName = "Calibri";
                    effect2.FontSize = 14;
                    // Underline not supported directly; can be handled via other formatting if needed
                }

                // Save the workbook with the modified WordArt shapes
                workbook.Save("WordArtUngroupedModified.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
