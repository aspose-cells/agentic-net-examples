// Title: Process SmartArt shapes only with Aspose.Cells for .NET
// Description: Load an Excel workbook, iterate every worksheet, filter shapes where IsSmartArt is true, convert each SmartArt to a GroupShape via GetResultOfSmartArt, shift the group 10 px right and down, and save the file.
// Keywords: Aspose.Cells SmartArt filter | C# GetResultOfSmartArt | convert SmartArt to GroupShape | move Excel shape programmatically | Aspose.Cells shape iteration | SmartArt positioning .NET
// Common Searches: Aspose.Cells only SmartArt shapes | C# convert SmartArt to grouped shape | how to move SmartArt with Aspose.Cells | filter worksheet shapes by IsSmartArt | batch adjust SmartArt position Excel
// Developer Intent: Identify SmartArt objects, turn them into GroupShape instances, and adjust their coordinates.
// Use Cases: Prevent overlapping SmartArt in generated reports by nudging each object. | Replace SmartArt with static groups for compatibility with older Excel versions. | Apply a uniform offset to all SmartArt across multiple sheets in a template.
// AI Prompts: Write C# code that finds every SmartArt shape in an Aspose.Cells workbook and changes its fill color. | Create a reusable method that extracts each SmartArt shape as a PNG using Aspose.Cells. | Generate a script that converts SmartArt to GroupShape and aligns them to a 50‑pixel grid.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load an Excel workbook, iterate every worksheet, filter shapes where IsSmartArt is true, convert each SmartArt to a GroupShape via GetResultOfSmartArt, shift the group 10 px right and down, and save the file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Loop through all shapes on the current worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only shapes that are SmartArt
                if (shape.IsSmartArt)
                {
                    // Convert the SmartArt shape to a grouped shape
                    GroupShape groupShape = shape.GetResultOfSmartArt();

                    // Example modification: shift the grouped shape slightly
                    if (groupShape != null)
                    {
                        groupShape.Left += 10;   // move 10 pixels to the right
                        groupShape.Top += 10;    // move 10 pixels down
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
