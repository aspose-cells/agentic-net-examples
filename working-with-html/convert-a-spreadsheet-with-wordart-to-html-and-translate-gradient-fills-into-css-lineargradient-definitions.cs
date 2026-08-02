// Title: Export Excel WordArt with Gradient Fills to HTML – CSS linear‑gradient via Aspose.Cells (C#)
// Description: Load an Excel workbook containing WordArt, set HtmlSaveOptions (ExportWorksheetCSSSeparately, EnableCssCustomProperties, HtmlVersion.Html5) and save it as HTML. The gradient fills are emitted as CSS linear‑gradient rules for modern browsers.
// Keywords: Aspose.Cells | C# export Excel to HTML | WordArt gradient to CSS | linear-gradient CSS | ExportWorksheetCSSSeparately | EnableCssCustomProperties | HTML5 Excel conversion | Excel to web report
// Common Searches: Aspose.Cells export WordArt gradient to HTML | HTML5 save options for Excel gradients | Convert Excel gradient fill to CSS linear‑gradient | ExportWorksheetCSSSeparately example C# | EnableCssCustomProperties Aspose.Cells
// Developer Intent: Create an HTML file from an Excel sheet that preserves WordArt gradient styling as CSS linear‑gradient definitions.
// Use Cases: Generate web‑ready reports from design‑heavy spreadsheets. | Automate publishing of Excel dashboards with accurate visual fidelity. | Integrate server‑side Excel‑to‑HTML conversion that relies on CSS‑only gradients.
// AI Prompts: Show how to inline the generated CSS into the HTML output instead of separate files. | Provide code to export each worksheet with WordArt gradients to its own HTML file. | Explain how to modify the linear‑gradient direction or colors after export.

using System;
using System.IO;
using Aspose.Cells;

// Load an Excel workbook containing WordArt, set HtmlSaveOptions (ExportWorksheetCSSSeparately, EnableCssCustomProperties, HtmlVersion.Html5) and save it as HTML. The gradient fills are emitted as CSS linear‑gradient rules for modern browsers.
class Program
{
    static void Main()
    {
        // Path to the source Excel file that contains WordArt with gradient fills
        string inputPath = "WordArtWithGradient.xlsx";

        // Desired output HTML file path
        string outputPath = "WordArtWithGradient.html";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (lifecycle: create)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export worksheet CSS separately so that gradient fills are emitted as CSS linear‑gradient definitions
                ExportWorksheetCSSSeparately = true,

                // Enable CSS custom properties to avoid duplicate CSS (optional but improves output)
                EnableCssCustomProperties = true,

                // Use HTML5 to ensure modern CSS features like linear‑gradient are supported
                HtmlVersion = HtmlVersion.Html5
            };

            // Save the workbook as HTML (lifecycle: save)
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
