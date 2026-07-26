// Title: Convert XLSX to HTML with Aspose.Cells (C#) – default settings preserve all content
// Description: Loads an XLSX workbook, applies default HtmlSaveOptions, and saves it as an HTML file, keeping formulas, images, charts, and formatting intact.
// Keywords: Aspose.Cells | C# Excel to HTML | XLSX to HTML conversion | HtmlSaveOptions default | preserve Excel content | export workbook as HTML | Aspose.Cells example | convert Excel to web page
// Common Searches: C# Aspose.Cells convert XLSX to HTML | Export Excel workbook to HTML default options | How to keep formulas and images when saving Excel as HTML | Aspose.Cells HtmlSaveOptions default behavior | Save Excel file as HTML with full content
// Developer Intent: Generate an HTML representation of an existing XLSX workbook using Aspose.Cells without altering any worksheet data or visual elements.
// Use Cases: Display Excel reports in browsers without requiring Office | Create static HTML snapshots of financial models for offline review | Provide preview of uploaded Excel files in web portals | Archive Excel worksheets as web‑compatible pages
// AI Prompts: Write C# code that reads an XLSX file and saves it as HTML using Aspose.Cells with default options, ensuring all content is retained. | Explain which default HtmlSaveOptions settings preserve formulas, images, charts, and styles during conversion. | Show how to change the output path or file name while still using default HTML conversion settings.

using System;
using Aspose.Cells;

// Loads an XLSX workbook, applies default HtmlSaveOptions, and saves it as an HTML file, keeping formulas, images, charts, and formatting intact.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Create default HTML save options (all default settings preserve full content)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export the workbook to an HTML file using the default options
        workbook.Save("output.html", htmlOptions);
    }
}
