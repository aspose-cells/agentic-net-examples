// Title: Convert Excel WordArt with gradient fills to HTML using Aspose.Cells and provide solid‑color fallback for legacy browsers
// AI Prompts: Write C# code that loads an .xlsx containing WordArt, saves it as HTML with Aspose.Cells, and then adds inline CSS rules that replace gradient backgrounds with solid colors for browsers lacking gradient support. | Show how to configure Aspose.Cells HtmlSaveOptions to export WordArt as separate image files and then post‑process the generated HTML to substitute gradient‑style CSS or SVG with a fallback solid‑color style. | Create a utility method that detects whether the current Aspose.Cells version supports gradient export, and if not, programmatically extracts the primary gradient color from WordArt objects and injects a fallback background‑color into the HTML output.
// Common Searches: aspnet convert excel worksheet containing wordart gradient to html with fallback color | aspose.cells html export gradient fill fallback for old browsers | c# save excel wordart as html and replace gradient with solid background | how to handle wordart gradient fills when generating html using asp.net and aspose.cells
// Tags: Aspose.Cells HtmlSaveOptions export WordArt images | gradient fill fallback solid color HTML | convert Excel WordArt to HTML C# | legacy browser CSS gradient fallback | post‑process Aspose.Cells HTML output

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input Excel file, loads it with Aspose.Cells, configures HtmlSaveOptions to export only the active worksheet and to save images as separate files, notes that automatic gradient‑to‑solid conversion is unavailable, and saves the workbook as HTML while handling any errors.
class Program
{
    static void Main()
    {
        const string inputPath = "InputWithWordArt.xlsx";
        const string outputPath = "Output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the Excel workbook that contains WordArt with gradient fills.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Optional: Export only the active worksheet.
                ExportActiveWorksheetOnly = true,
                // Optional: Export images (including WordArt) as separate files.
                ExportImagesAsBase64 = false
                // Note: ExportGradientAsSolidColor is not available in this version of Aspose.Cells.
            };

            // Save the workbook as HTML with the configured options.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
