// Title: Batch convert Excel workbooks to HTML with per‑file default fonts using Aspose.Cells for .NET
// Description: This C# example shows how to map each Excel file to a specific default font, load the workbook with Aspose.Cells, configure HtmlSaveOptions.DefaultFontName, and save the result as an HTML page. The loop processes a dictionary of file‑font pairs, producing individually styled HTML outputs in a single run.
// Keywords: Aspose.Cells batch HTML export | HtmlSaveOptions DefaultFontName | C# convert Excel to HTML | per workbook custom font | multiple Excel to HTML .NET | automated spreadsheet conversion | Aspose.Cells HTML conversion example | custom font HTML export | dictionary driven file processing | Excel to web‑ready HTML
// Common Searches: Aspose.Cells batch convert Excel to HTML with different fonts | C# set default font for each HTML export using Aspose | How to use HtmlSaveOptions.DefaultFontName in a loop | Convert multiple .xlsx files to .html with custom fonts | Aspose.Cells example for per‑file HTML styling
// Developer Intent: Generate HTML files from a set of Excel workbooks, applying a distinct default font to each output in a single automated process.
// Use Cases: Produce branded HTML reports where each report follows a different corporate typeface. | Offer a web service that converts user‑uploaded spreadsheets to HTML, respecting the user’s preferred font. | Run a nightly job that archives legacy Excel files as HTML, assigning readable fonts based on file categories.
// AI Prompts: Write a C# script that reads a JSON array of {filePath, fontName} objects and converts each workbook to HTML with Aspose.Cells, using HtmlSaveOptions.DefaultFontName. | Explain how to modify the batch conversion to store all HTML files in a dedicated output folder while preserving the per‑file font settings. | Suggest error‑handling strategies for missing or unsupported font names during batch HTML export with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// This C# example shows how to map each Excel file to a specific default font, load the workbook with Aspose.Cells, configure HtmlSaveOptions.DefaultFontName, and save the result as an HTML page. The loop processes a dictionary of file‑font pairs, producing individually styled HTML outputs in a single run.
class BatchHtmlConversion
{
    static void Main()
    {
        // Define source Excel files and the default font to use for each HTML output
        var filesAndFonts = new Dictionary<string, string>
        {
            { "Book1.xlsx", "Arial" },
            { "Book2.xlsx", "Courier New" },
            { "Book3.xlsx", "Times New Roman" }
        };

        foreach (var entry in filesAndFonts)
        {
            string sourcePath = entry.Key;      // Path to the Excel workbook
            string defaultFont = entry.Value;   // Desired default font for HTML

            // Load the workbook (load lifecycle)
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options and set the distinct default font (create lifecycle)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.DefaultFontName = defaultFont;

            // Build the output HTML file name
            string outputPath = System.IO.Path.ChangeExtension(sourcePath, ".html");

            // Save the workbook as HTML using the configured options (save lifecycle)
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Converted '{sourcePath}' to HTML with default font '{defaultFont}'.");
        }
    }
}
