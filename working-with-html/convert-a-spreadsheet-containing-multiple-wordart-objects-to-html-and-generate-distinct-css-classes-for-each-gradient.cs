// Title: Export Excel WordArt to HTML with Separate Gradient CSS Classes using Aspose.Cells for .NET
// Description: Loads an Excel workbook containing multiple WordArt shapes, configures HtmlSaveOptions to export worksheet CSS separately and retain all styles, then saves the file as HTML where each WordArt gradient is emitted as a unique CSS class.
// Keywords: Aspose.Cells | C# | WordArt to HTML | gradient CSS class | ExportWorksheetCSSSeparately | HtmlSaveOptions | Excel to web | preserve WordArt styling
// Common Searches: Aspose.Cells export WordArt gradients as separate CSS classes | HTML save options for WordArt in Excel .NET | how to keep all gradient styles when converting Excel to HTML | ExportWorksheetCSSSeparately example C# | convert Excel WordArt to web‑ready HTML
// Developer Intent: Create an HTML representation of an Excel file that contains several WordArt objects, ensuring each gradient used by a WordArt shape is generated as its own CSS class.
// Use Cases: Publish Excel dashboards with accurate WordArt gradients on corporate intranets. | Generate web‑friendly reports where each worksheet’s WordArt styling can be customized independently. | Produce HTML email templates that retain the original WordArt visual effects.
// AI Prompts: Show how to rename the generated gradient CSS classes based on the WordArt shape name. | Provide code to extract the CSS file created by HtmlSaveOptions and embed it into a custom HTML layout. | Explain how to disable ExportWorksheetCSSSeparately while still outputting all gradient styles in a single stylesheet.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook containing multiple WordArt shapes, configures HtmlSaveOptions to export worksheet CSS separately and retain all styles, then saves the file as HTML where each WordArt gradient is emitted as a unique CSS class.
class WordArtToHtml
{
    static void Main()
    {
        // Load the workbook that contains multiple WordArt objects.
        string inputFile = "WordArtSample.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Configure HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Export CSS for each worksheet separately so that each gradient
        // used by a WordArt shape gets its own CSS class.
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Keep all generated styles (including gradient classes) in the output.
        saveOptions.ExcludeUnusedStyles = false;

        // Ensure the output directory exists.
        string outputDir = "HtmlOutput";
        Directory.CreateDirectory(outputDir);

        // Save the workbook as an HTML file.
        string htmlFile = Path.Combine(outputDir, "WordArt.html");
        workbook.Save(htmlFile, saveOptions);

        Console.WriteLine($"HTML file with distinct gradient CSS classes saved to: {htmlFile}");
    }
}
