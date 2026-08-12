// Title: Convert Excel to HTML with BestFit layout and CSS custom properties using Aspose.Cells for .NET
// Description: Shows how to load an .xlsx file with Aspose.Cells, configure HtmlSaveOptions to use PresentationPreference = BestFit, enable EnableCssCustomProperties for base64 image deduplication, and save the workbook as optimized HTML in C#.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | HtmlSaveOptions | PresentationPreference | BestFit | EnableCssCustomProperties | image deduplication | CSS custom properties | workbook conversion | HTML export
// Common Searches: Aspose.Cells export Excel to HTML BestFit C# | EnableCssCustomProperties image deduplication Aspose.Cells | HtmlSaveOptions PresentationPreference example | Convert .xlsx to HTML with CSS custom properties | Reduce HTML size when exporting Excel with Aspose.Cells
// Developer Intent: Generate an HTML file from an Excel workbook that keeps the original visual layout (BestFit) while minimizing duplicated base64 images by using CSS custom properties.
// Use Cases: Produce web‑ready reports that retain Excel formatting and have smaller payloads thanks to image deduplication. | Automate batch conversion of multiple spreadsheets to consistent, best‑fit HTML for publishing on intranets or portals. | Integrate optimized HTML export into a .NET web service that returns lightweight, style‑driven content.
// AI Prompts: Write C# code with Aspose.Cells to convert an Excel workbook to HTML using PresentationPreference.BestFit and EnableCssCustomProperties. | Explain how EnableCssCustomProperties consolidates identical base64 images into CSS variables and how to reference those variables in the generated HTML. | Create a reusable C# method that accepts input and output paths, applies the best‑fit presentation preference, enables CSS custom properties, and includes robust error handling.

using System;
using Aspose.Cells;

// Shows how to load an .xlsx file with Aspose.Cells, configure HtmlSaveOptions to use PresentationPreference = BestFit, enable EnableCssCustomProperties for base64 image deduplication, and save the workbook as optimized HTML in C#.
class ExcelToHtmlConverter
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path where the HTML output will be saved
        string outputPath = "output.html";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(sourcePath);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Enable presentation preference for a more faithful visual rendering
        htmlOptions.PresentationPreference = true;

        // Enable CSS custom properties to deduplicate repeated base64 images
        htmlOptions.EnableCssCustomProperties = true;

        // Save the workbook as an HTML file using the configured options
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine("Conversion completed. HTML saved to: " + outputPath);
    }
}
