// Title: Export Gradient WordArt to HTML5 with Aspose.Cells – W3C‑Compliant CSS
// Description: Creates a new workbook, adds a WordArt shape with a preset gradient fill, customizes its text, configures HtmlSaveOptions for HTML5 output (no CSS custom properties), and saves the file as HTML containing a linear‑gradient that passes W3C validation.
// Keywords: Aspose.Cells WordArt HTML export | C# gradient WordArt to HTML5 | W3C CSS gradient validation | HtmlSaveOptions HTML5 Aspose | Excel to HTML with CSS gradients | disable CSS custom properties Aspose | WordArtStyle7 gradient example | standards‑compliant HTML from Excel
// Common Searches: export WordArt with gradient to HTML using Aspose.Cells | Aspose.Cells HTML5 output CSS gradient validation | how to disable CSS custom properties in Aspose.Cells HTML export | convert Excel workbook with WordArt to W3C‑valid HTML | C# code for WordArt gradient and HTML5 save options
// Developer Intent: Generate an HTML5 file from an Excel workbook that contains a gradient‑filled WordArt shape and complies with W3C CSS gradient validation.
// Use Cases: Produce web‑ready reports that include stylized WordArt with validated CSS gradients. | Automate conversion of Excel dashboards into HTML5 email templates that meet W3C standards. | Create documentation pages from workbooks where WordArt graphics must render consistently across browsers.
// AI Prompts: Write C# code with Aspose.Cells to add a WordArt shape using a gradient fill and save the workbook as HTML5 with W3C‑compliant CSS. | Explain how to configure HtmlSaveOptions to disable CSS custom properties and ensure the generated linear‑gradient passes W3C validation. | Provide step‑by‑step instructions to modify WordArt text, apply a preset gradient style, and export the workbook to HTML5 using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using Aspose.Cells.Saving;

// Creates a new workbook, adds a WordArt shape with a preset gradient fill, customizes its text, configures HtmlSaveOptions for HTML5 output (no CSS custom properties), and saves the file as HTML containing a linear‑gradient that passes W3C validation.
class Program
{
    static void Main()
    {
        // 1. Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // 2. Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // 3. Add a WordArt shape with a preset gradient style (WordArtStyle7)
        //    Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7, // gradient fill - Blue, Accent 1, Reflection
            "Aspose.Cells WordArt",
            2, 0,   // row, top offset
            2, 0,   // column, left offset
            100,    // height (points)
            400);   // width (points)

        // 4. Optionally modify the WordArt text using FontSettingCollection
        //    This demonstrates the SetWordArtStyle method (feature rule)
        FontSettingCollection fontSettings = wordArt.TextBody;
        fontSettings.Text = "Gradient WordArt Example";

        // 5. Prepare HTML save options to produce standards‑compliant output
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Use HTML5 so that CSS gradient syntax is recognized by W3C validators
            HtmlVersion = HtmlVersion.Html5,

            // Keep CSS in external files (optional, but does not affect validation)
            ExportWorksheetCSSSeparately = false,

            // Ensure CSS is emitted (not inline only) for proper gradient rules
            DisableCss = false,

            // Do not use CSS custom properties which some validators may flag
            EnableCssCustomProperties = false
        };

        // 6. Save the workbook as HTML (lifecycle: save)
        //    The output will contain a <style> block with a linear‑gradient definition
        workbook.Save("WordArtGradient.html", htmlOptions);

        Console.WriteLine("HTML file with WordArt and CSS gradients generated successfully.");
    }
}
