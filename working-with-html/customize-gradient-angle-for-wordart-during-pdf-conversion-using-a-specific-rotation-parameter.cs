using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class GradientWordArtPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset gradient style
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
            "Gradient WordArt",
            2,   // topRow
            0,   // top (pixels)
            2,   // leftColumn
            0,   // left (pixels)
            200, // height (pixels)
            100  // width (pixels)
        );

        // Ensure the fill type is gradient
        wordArt.Fill.FillType = FillType.Gradient;

        // Set a linear gradient (angle will be overridden by the Angle property)
        wordArt.Fill.GradientFill.SetGradient(GradientFillType.Linear, 0, GradientDirectionType.FromCenter);

        // Customize the gradient angle (e.g., 30 degrees)
        wordArt.Fill.GradientFill.Angle = 30f;

        // Define the gradient colors
        wordArt.Fill.GradientFill.SetTwoColorGradient(
            Color.Blue,
            Color.LightBlue,
            GradientStyleType.Horizontal,
            1
        );

        // Save the workbook as PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("GradientWordArt.pdf", pdfOptions);
    }
}