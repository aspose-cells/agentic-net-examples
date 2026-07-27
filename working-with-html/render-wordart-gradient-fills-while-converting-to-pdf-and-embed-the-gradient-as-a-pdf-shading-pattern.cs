// Title: Add Gradient WordArt and Export to PDF with Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a WordArt shape using the preset WordArtStyle7 (blue gradient), and save the file as PDF. Aspose.Cells automatically converts the gradient into a PDF shading pattern, preserving the visual effect.
// Keywords: Aspose.Cells | .NET | WordArt | gradient fill | PDF export | shading pattern | PresetWordArtStyle | AddWordArt | SaveFormat.Pdf | Excel to PDF
// Common Searches: How to add gradient WordArt in Excel with Aspose.Cells | Export WordArt with gradient to PDF using Aspose.Cells .NET | Does Aspose.Cells preserve WordArt gradients in PDF | PresetWordArtStyle gradient options in Aspose.Cells | Customize WordArt gradient colors before PDF conversion
// Developer Intent: Create a WordArt shape with a gradient fill and ensure the gradient is retained when the workbook is saved as a PDF.
// Use Cases: Design report headers with gradient WordArt that appear correctly in PDF output. | Generate marketing flyers where decorative gradient WordArt must be preserved in the final PDF. | Programmatically adjust WordArt gradient colors before exporting an Excel workbook to PDF.
// AI Prompts: Provide code to change the two‑color gradient of a WordArt shape before saving to PDF with Aspose.Cells for .NET. | Explain how Aspose.Cells maps WordArt gradient fills to PDF shading patterns during the Save operation. | Show an example of applying a custom horizontal gradient to WordArt and exporting it as a PDF with the gradient embedded.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to create a workbook, insert a WordArt shape using the preset WordArtStyle7 (blue gradient), and save the file as PDF. Aspose.Cells automatically converts the gradient into a PDF shading pattern, preserving the visual effect.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape using a preset style that includes a gradient fill
        // WordArtStyle7 = Gradient Fill - Blue, Accent 1, Reflection
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2,          // upperLeftRow
            0,          // top
            2,          // upperLeftColumn
            0,          // left
            200,        // height
            400);       // width

        // The preset style already applies the gradient; additional customization can be done if needed:
        // wordArt.Fill.FillType = FillType.Gradient;
        // wordArt.Fill.SetTwoColorGradient(Color.Blue, Color.LightBlue, GradientStyleType.Horizontal, 1);

        // Save the workbook as PDF (lifecycle save)
        // Aspose.Cells automatically embeds the gradient as a PDF shading pattern.
        workbook.Save("WordArtGradient.pdf", SaveFormat.Pdf);
    }
}
