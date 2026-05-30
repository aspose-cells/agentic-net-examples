using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Store custom start and end colors as ARGB integers in cells
            // Use unchecked to allow values larger than Int32.MaxValue
            sheet.Cells["A1"].PutValue(unchecked((int)0xFFFF0000)); // Red (ARGB)
            sheet.Cells["B1"].PutValue(unchecked((int)0xFF0000FF)); // Blue (ARGB)

            // Retrieve the colors from the cells as integers
            int startArgb = sheet.Cells["A1"].IntValue;
            int endArgb   = sheet.Cells["B1"].IntValue;

            Color startColor = Color.FromArgb(startArgb);
            Color endColor   = Color.FromArgb(endArgb);

            // Add a WordArt shape (any preset style can be used as a base)
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // base preset
                "Custom Gradient WordArt",        // text
                2, 0, 2, 0,                       // position: upperLeftRow, top, upperLeftColumn, left
                300, 100);                        // size: height, width

            // Set the fill type to gradient
            wordArt.Fill.FillType = FillType.Gradient;

            // Apply a two‑color gradient using the custom start and end colors
            wordArt.Fill.SetTwoColorGradient(
                startColor,                     // custom start color
                endColor,                       // custom end color
                GradientStyleType.Horizontal,   // gradient direction
                1);                             // variant (1‑4)

            // Save the workbook with the WordArt shape rendered with the custom gradient
            workbook.Save("WordArtCustomGradient.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}