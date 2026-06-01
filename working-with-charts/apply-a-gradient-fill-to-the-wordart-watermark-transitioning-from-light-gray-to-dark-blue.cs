using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class GradientWordArtWatermark
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape that will act as a watermark
            // Parameters: preset style, text, upperLeftRow, top, upperLeftColumn, left, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "CONFIDENTIAL",
                5, 0,   // row and vertical offset
                5, 0,   // column and horizontal offset
                300, 100); // height and width

            // Position and size the shape (centered roughly)
            wordArt.Left = 0;
            wordArt.Top = 0;
            wordArt.Width = 500;
            wordArt.Height = 100;

            // Apply a two‑color gradient fill: light gray to dark blue
            wordArt.Fill.FillType = FillType.Gradient;
            GradientFill gradient = wordArt.Fill.GradientFill;
            gradient.SetTwoColorGradient(
                Color.LightGray,
                Color.DarkBlue,
                GradientStyleType.Horizontal,
                1);

            // Adjust text formatting (font, size, bold)
            if (wordArt.IsWordArt)
            {
                TextEffectFormat txtEffect = wordArt.TextEffect;
                txtEffect.FontName = "Arial";
                txtEffect.FontSize = 48;
                txtEffect.FontBold = true;
                txtEffect.FontItalic = false;
            }

            // Ensure the output directory exists
            string outputPath = "GradientWordArtWatermark.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}