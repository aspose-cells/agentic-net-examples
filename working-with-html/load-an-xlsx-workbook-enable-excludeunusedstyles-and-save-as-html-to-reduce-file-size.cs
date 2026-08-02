// Title: Convert XLSX to lightweight HTML with Aspose.Cells for .NET (ExcludeUnusedStyles)
// Description: Loads an XLSX workbook, enables HtmlSaveOptions.ExcludeUnusedStyles to omit unused CSS, and saves the file as a compact HTML document, reducing output size.
// Keywords: Aspose.Cells | C# | .NET | HTML export | ExcludeUnusedStyles | optimize HTML size | Excel to HTML conversion | reduce HTML payload | Aspose.Cells HtmlSaveOptions | lightweight HTML report
// Common Searches: Aspose.Cells exclude unused styles HTML example | How to shrink HTML output when converting Excel with Aspose | C# convert XLSX to HTML without extra CSS | HtmlSaveOptions ExcludeUnusedStyles usage | Optimize Excel to HTML conversion Aspose.Cells
// Developer Intent: Export an Excel workbook to HTML while automatically removing any style definitions that are not used, to produce a smaller, faster‑loading file.
// Use Cases: Create fast‑loading web reports from Excel data. | Batch‑process workbooks for email‑friendly HTML attachments. | Serve Excel‑derived content over low‑bandwidth connections. | Archive spreadsheets as minimal‑size HTML snapshots.
// AI Prompts: Generate C# code that loads an XLSX file with Aspose.Cells and saves it as HTML using ExcludeUnusedStyles, and explain the size benefit. | Describe how HtmlSaveOptions.ExcludeUnusedStyles affects the CSS generated for a workbook containing many custom styles. | Show a combined Aspose.Cells HTML export configuration that uses ExcludeUnusedStyles together with ExportImagesAsBase64 for the smallest possible HTML output.

using System;
using Aspose.Cells;

// Loads an XLSX workbook, enables HtmlSaveOptions.ExcludeUnusedStyles to omit unused CSS, and saves the file as a compact HTML document, reducing output size.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable exclusion of unused styles
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExcludeUnusedStyles = true; // Reduces HTML file size

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", saveOptions);
    }
}
