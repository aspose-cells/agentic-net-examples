// Title: Convert Excel WordArt to HTML with individual CSS gradient classes using Aspose.Cells (.NET)
// Description: Loads an .xlsx, renames each WordArt shape, configures HtmlSaveOptions (ExportWorksheetCSSSeparately = true, ExcludeUnusedStyles = false) and saves HTML with separate CSS files so every gradient gets its own class.
// Keywords: Aspose.Cells | C# | WordArt | HTML export | ExportWorksheetCSSSeparately | ExcludeUnusedStyles | gradient CSS | Excel to HTML | .NET | shape naming
// Common Searches: Aspose.Cells export WordArt to HTML | HTML save options for WordArt gradients | How to generate separate CSS for each WordArt shape | C# convert Excel WordArt to HTML | Preserve WordArt gradient styles in HTML
// Developer Intent: Generate HTML from an Excel workbook that preserves each WordArt object's gradient in a unique CSS class.
// Use Cases: Publish Excel reports with WordArt titles on a website while keeping exact gradient styling. | Batch‑convert multiple workbooks containing WordArt into HTML/CSS for intranet documentation. | Create a branding tool that extracts WordArt, assigns IDs, and outputs HTML so corporate gradients stay consistent.
// AI Prompts: Write C# code with Aspose.Cells to convert an Excel file containing several WordArt objects to HTML, assigning unique names to each shape and exporting CSS separately for distinct gradient classes. | Explain how ExportWorksheetCSSSeparately and ExcludeUnusedStyles affect the CSS generated for WordArt‑rich worksheets when saving to HTML with Aspose.Cells. | Provide a step‑by‑step guide to batch process a folder of .xlsx files that contain WordArt, producing HTML and separate CSS files where each gradient has its own class.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an .xlsx, renames each WordArt shape, configures HtmlSaveOptions (ExportWorksheetCSSSeparately = true, ExcludeUnusedStyles = false) and saves HTML with separate CSS files so every gradient gets its own class.
class ConvertWordArtToHtml
{
    static void Main()
    {
        // Load the workbook that contains multiple WordArt objects
        Workbook workbook = new Workbook("WordArtSample.xlsx");

        // Assign a unique name to each WordArt shape.
        // The shape name is used by Aspose.Cells to generate distinct CSS class names.
        Worksheet worksheet = workbook.Worksheets[0];
        ShapeCollection shapes = worksheet.Shapes;
        int wordArtIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.IsWordArt)
            {
                shape.Name = $"WordArt_{++wordArtIndex}";
            }
        }

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export the worksheet CSS to a separate file so that each gradient gets its own CSS class.
        htmlOptions.ExportWorksheetCSSSeparately = true;

        // Keep all generated styles (do not exclude unused ones) to ensure each gradient style is retained.
        htmlOptions.ExcludeUnusedStyles = false;

        // Specify a directory where the HTML file and its associated CSS will be saved.
        htmlOptions.AttachedFilesDirectory = "HtmlOutput";
        htmlOptions.CreateDirectory = true;

        // Save the workbook as HTML using the configured options.
        workbook.Save("HtmlOutput/WordArt.html", htmlOptions);
    }
}
