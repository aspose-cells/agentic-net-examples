// Title: Add a Semi‑Transparent WordArt Watermark to Every Worksheet with Aspose.Cells (C#)
// Description: Demonstrates how to loop through all worksheets in a workbook, insert a WordArt shape with custom text, apply 50 % fill transparency, hide its outline, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# watermark | WordArt transparency Aspose | add watermark all sheets | Excel WordArt shape | iterate worksheets Aspose.Cells | semi transparent watermark C# | hide WordArt outline | Aspose.Cells shape collection
// Common Searches: Aspose.Cells add watermark to each worksheet | C# WordArt watermark transparency Excel | How to loop through worksheets and insert WordArt with Aspose.Cells | Set WordArt fill transparency in Aspose.Cells | Hide WordArt border in Excel using Aspose
// Developer Intent: Insert a semi‑transparent WordArt overlay on every worksheet of an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Protect confidential reports by overlaying a faint "CONFIDENTIAL" label on all sheets. | Brand internal templates with a company slogan or logo as a subtle background element. | Add a legal disclaimer or copyright notice across an entire workbook before distribution.
// AI Prompts: Write C# code that uses Aspose.Cells to add a WordArt watermark with customizable text, size, position, and 50 % transparency to each worksheet. | Show how to adjust WordArt fill transparency and hide its outline when creating a watermark in an Excel file with Aspose.Cells. | Provide a complete example that iterates over all worksheets, inserts a semi‑transparent WordArt shape, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to loop through all worksheets in a workbook, insert a WordArt shape with custom text, apply 50 % fill transparency, hide its outline, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty with a default worksheet)
        Workbook workbook = new Workbook();

        // Optional: add some sample data to each sheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.Cells["A1"].PutValue("Sample data");
        }

        // Iterate through all worksheets and add a semi‑transparent WordArt watermark
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Get the shape collection of the current worksheet
            ShapeCollection shapes = ws.Shapes;

            // Add WordArt. Parameters: style, text, topRow, top offset, leftColumn, left offset, height, width
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "CONFIDENTIAL",
                0,    // topRow index
                0,    // top offset (pixels)
                0,    // leftColumn index
                0,    // left offset (pixels)
                100,  // height (pixels)
                400   // width (pixels)
            );

            // Set semi‑transparent fill (0 = opaque, 1 = fully transparent)
            wordArt.FillFormat.Transparency = 0.5;

            // Hide the outline of the WordArt
            wordArt.LineFormat.IsVisible = false;
        }

        // Save the workbook with the watermarks applied
        workbook.Save("WorkbookWithWatermark.xlsx");
    }
}
