using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample worksheets (some start with "Report")
        workbook.Worksheets[0].Name = "ReportJan";
        workbook.Worksheets.Add("ReportFeb");
        workbook.Worksheets.Add("Data");

        // Add a WordArt shape to each worksheet for demonstration purposes
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // any preset style
                "Sample Text",
                1,   // topRow
                0,   // top (pixels)
                1,   // leftColumn
                0,   // left (pixels)
                100, // height (pixels)
                400  // width (pixels)
            );
        }

        // Apply the "Bold Wave" style to all WordArt shapes on worksheets whose names start with "Report"
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
            {
                foreach (Shape shape in ws.Shapes)
                {
                    if (shape.IsWordArt)
                    {
                        // Make the text bold
                        shape.TextEffect.FontBold = true;
                        // Set the preset shape to Wave1 (represents a wave effect)
                        shape.TextEffect.PresetShape = MsoPresetTextEffectShape.Wave1;
                    }
                }
            }
        }

        // Save the workbook
        workbook.Save("ReportWordArtBoldWave.xlsx");
    }
}