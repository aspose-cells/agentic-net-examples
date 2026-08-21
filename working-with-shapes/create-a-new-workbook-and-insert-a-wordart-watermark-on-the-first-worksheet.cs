// Title: Add a WordArt Watermark to a New Excel Workbook with Aspose.Cells (C#)
// Description: Creates a fresh Workbook, accesses the first worksheet's ShapeCollection, inserts a WordArt shape with custom text as a watermark, and saves the file as WordArtWatermark.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# WordArt watermark | add WordArt shape Excel | Excel watermark Aspose.Cells | create workbook with watermark | ShapeCollection AddWordArt | C# Excel watermark example
// Common Searches: how to add a WordArt watermark in Excel using Aspose.Cells | Aspose.Cells C# insert WordArt on first worksheet | example code for Excel watermark with Aspose.Cells | add confidential text overlay to Excel file C#
// Developer Intent: Insert a WordArt shape as a watermark on the first sheet of a newly created workbook.
// Use Cases: Mark confidential reports with a visible "CONFIDENTIAL" overlay. | Brand internal templates by adding a company slogan as WordArt. | Enhance printable spreadsheets with decorative text for emphasis.
// AI Prompts: Show how to change the color and transparency of the WordArt watermark in Aspose.Cells. | Provide code to send the WordArt watermark behind cell content and lock it from editing. | Explain how to apply the same WordArt watermark to all worksheets in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a fresh Workbook, accesses the first worksheet's ShapeCollection, inserts a WordArt shape with custom text as a watermark, and saves the file as WordArtWatermark.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the shape collection of the first worksheet
        ShapeCollection shapes = sheet.Shapes;

        // Add WordArt as a watermark
        // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle2, // preset style
            "CONFIDENTIAL",                  // watermark text
            0,                               // top row index
            0,                               // top offset in pixels
            0,                               // left column index
            0,                               // left offset in pixels
            100,                             // height in pixels
            400                              // width in pixels
        );

        // Save the workbook with the WordArt watermark
        workbook.Save("WordArtWatermark.xlsx");
    }
}
