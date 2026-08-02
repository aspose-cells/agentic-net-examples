// Title: Export Aspose.Cells Workbook with Gradient WordArt to HTML5 – W3C‑Compliant CSS
// Description: Shows how to insert a WordArt shape (PresetWordArtStyle.WordArtStyle7) into a workbook, set HtmlSaveOptions (HtmlVersion=Html5, ExportWorksheetCSSSeparately=true, ExcludeUnusedStyles=false) and save as HTML5 with external CSS so the gradient fill remains intact and validates against the W3C CSS validator.
// Keywords: Aspose.Cells | C# export WordArt to HTML | HTML5 | CSS gradient validation | HtmlSaveOptions | PresetWordArtStyle | WordArtStyle7 | external CSS file | W3C validator | convert workbook to HTML
// Common Searches: Aspose.Cells export WordArt gradient to HTML5 | HTML5 output with CSS gradients from Aspose.Cells | How to keep WordArt styles when saving as HTML in .NET | W3C CSS validation for Aspose.Cells HTML export | Separate CSS file for Aspose.Cells HTML5 conversion
// Developer Intent: Create an HTML5 document from a workbook that contains gradient‑filled WordArt, ensuring the generated CSS is external, complete, and passes W3C validation.
// Use Cases: Add a gradient WordArt shape to a worksheet and produce standards‑compliant HTML5 for web publishing. | Export multiple worksheets with WordArt to individual HTML files while preserving each sheet’s CSS gradients. | Integrate the exported HTML and CSS into a corporate portal that requires W3C‑certified markup.
// AI Prompts: Generate C# code that adds a WordArt shape with a gradient fill using Aspose.Cells and saves the workbook as HTML5 with external CSS for W3C validation. | Explain the impact of HtmlSaveOptions properties HtmlVersion, ExportWorksheetCSSSeparately, and ExcludeUnusedStyles on preserving WordArt gradients in the output HTML. | Provide troubleshooting steps when the CSS gradient is missing or fails W3C validation after exporting a workbook with WordArt.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to insert a WordArt shape (PresetWordArtStyle.WordArtStyle7) into a workbook, set HtmlSaveOptions (HtmlVersion=Html5, ExportWorksheetCSSSeparately=true, ExcludeUnusedStyles=false) and save as HTML5 with external CSS so the gradient fill remains intact and validates against the W3C CSS validator.
class ConvertWordArtToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a gradient preset style (WordArtStyle7)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
            "Aspose WordArt",
            2, 0,   // Row and top offset
            2, 0,   // Column and left offset
            100,    // Height
            400);   // Width

        // Configure HTML save options to produce standards‑compliant output
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Use HTML5 so that CSS gradients are recognized by W3C validators
        htmlOptions.HtmlVersion = HtmlVersion.Html5;

        // Export CSS to a separate file for easier validation (optional)
        htmlOptions.ExportWorksheetCSSSeparately = true;

        // Keep all generated styles so gradient definitions are not stripped out
        htmlOptions.ExcludeUnusedStyles = false;

        // Save the workbook as an HTML file
        workbook.Save("WordArtOutput.html", htmlOptions);
    }
}
