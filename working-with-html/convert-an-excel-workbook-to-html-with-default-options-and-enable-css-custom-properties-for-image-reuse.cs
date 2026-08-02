// Title: C# – Convert Excel to HTML with CSS custom properties (image deduplication) using Aspose.Cells
// Description: Loads an .xlsx workbook with default options, creates HtmlSaveOptions, enables EnableCssCustomProperties to store repeated base64 images once, and saves the spreadsheet as an HTML file.
// Keywords: Aspose.Cells | Excel to HTML conversion | EnableCssCustomProperties | CSS custom properties image reuse | C# Aspose.Cells HTML export | base64 image deduplication | default HtmlSaveOptions
// Common Searches: Aspose.Cells export Excel to HTML with CSS custom properties | EnableCssCustomProperties C# example | How to deduplicate images when saving Excel as HTML | Default HtmlSaveOptions Aspose.Cells | C# convert .xlsx to HTML with image reuse
// Developer Intent: Export an Excel workbook to HTML while activating CSS custom properties to eliminate duplicate base64 images.
// Use Cases: Create lightweight web previews of spreadsheets by reusing images via CSS variables. | Automate batch conversion of multiple .xlsx files to HTML with built‑in image optimization. | Integrate Excel‑to‑HTML conversion into a web application that requires minimal page size.
// AI Prompts: Generate C# code that converts an Excel file to HTML with EnableCssCustomProperties set to true using Aspose.Cells. | Explain how EnableCssCustomProperties reduces duplicate base64 images in the HTML output and how to reference the CSS variables. | Provide a script to batch process a folder of .xlsx files into HTML files with CSS custom properties enabled.

using System;
using Aspose.Cells;

// Loads an .xlsx workbook with default options, creates HtmlSaveOptions, enables EnableCssCustomProperties to store repeated base64 images once, and saves the spreadsheet as an HTML file.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file (default load options are used)
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options with default settings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Enable CSS custom properties so that repeated base64 images are stored once
        htmlOptions.EnableCssCustomProperties = true;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
