// Title: Convert Excel to Minimal HTML (Disable CSS) with Aspose.Cells for .NET
// Description: Loads an .xlsx workbook using Aspose.Cells, creates default HtmlSaveOptions, sets DisableCss = true to suppress stylesheet generation, and saves the result as inline‑styled HTML. Suitable for lightweight web pages or email bodies.
// Keywords: Aspose.Cells | Excel to HTML | DisableCss | HtmlSaveOptions | C# conversion | minimal HTML | inline styling | .NET | export spreadsheet to HTML | no CSS
// Common Searches: Aspose.Cells export Excel to HTML without CSS | C# HtmlSaveOptions DisableCss example | convert .xlsx to plain HTML .NET | how to disable CSS in Aspose.Cells HTML export | minimal HTML output from Excel
// Developer Intent: Generate an HTML file from an Excel workbook while omitting external CSS, producing a compact document with inline styles.
// Use Cases: Render spreadsheet data on web pages where external style sheets are prohibited. | Embed Excel content in email messages that require self‑contained HTML. | Create quick previews of workbooks in environments with strict CSS policies.
// AI Prompts: Write C# code that uses Aspose.Cells to convert an .xlsx file to HTML with the DisableCss option enabled. | Explain how the DisableCss property changes the HTML output of Aspose.Cells and suggest ways to style the result afterward. | Show how to export an Excel workbook to HTML with inline images (base64) while keeping CSS disabled using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an .xlsx workbook using Aspose.Cells, creates default HtmlSaveOptions, sets DisableCss = true to suppress stylesheet generation, and saves the result as inline‑styled HTML. Suitable for lightweight web pages or email bodies.
class ExcelToHtmlConverter
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path where the HTML output will be saved
        string outputPath = "output.html";

        // Load the workbook from the Excel file (uses the provided load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Create HTML save options with default settings (uses the provided constructor rule)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Disable CSS generation to produce minimal inline‑styled HTML
        saveOptions.DisableCss = true;

        // Save the workbook as HTML using the configured options (uses the provided save rule)
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Conversion completed. HTML saved to: " + outputPath);
    }
}
