// Title: Convert Aspose.Cells Workbook to XPS with printing‑optimized options (C#)
// Description: Creates a Workbook, adds sample data, configures XpsSaveOptions (OnePagePerSheet, DefaultFont, font compatibility checks, page range) and saves the file as an XPS document suitable for high‑quality printing.
// Keywords: Aspose.Cells | XPS conversion | C# | .NET | XpsSaveOptions | OnePagePerSheet | DefaultFont | font compatibility | export Excel to XPS | printable XPS
// Common Searches: Aspose.Cells export Excel to XPS C# | How to save workbook as XPS with one page per sheet | C# XpsSaveOptions example | Convert Excel to XPS for printing using Aspose | Set default font when converting to XPS Aspose.Cells
// Developer Intent: Generate an XPS file from an Excel workbook using Aspose.Cells with settings that improve print layout and font consistency.
// Use Cases: Produce a print‑ready XPS snapshot of a financial report. | Create a single‑page XPS preview for email attachment without sharing the full workbook. | Archive Excel data as XPS while preserving exact visual appearance across platforms. | Generate XPS files for batch printing of invoices directly from .NET applications.
// AI Prompts: Show how to export every worksheet to a separate XPS page with Aspose.Cells. | Provide code that customizes page size and margins in XpsSaveOptions before saving. | Explain how to disable font compatibility checks while keeping a specific default font during XPS conversion.

using System;
using Aspose.Cells;

// Creates a Workbook, adds sample data, configures XpsSaveOptions (OnePagePerSheet, DefaultFont, font compatibility checks, page range) and saves the file as an XPS document suitable for high‑quality printing.
public class ConvertWorkbookToXps
{
    public static void Run()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("XPS conversion demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Initialize XPS save options
            XpsSaveOptions xpsOptions = new XpsSaveOptions
            {
                // Render each sheet on a single page for better printing layout
                OnePagePerSheet = true,
                // Use a common font to ensure consistent rendering
                DefaultFont = "Arial",
                // Enable font compatibility checks for Unicode characters
                CheckFontCompatibility = true,
                CheckWorkbookDefaultFont = true,
                // Save only the first page (adjust as needed)
                PageIndex = 0,
                PageCount = 1
            };

            // Save the workbook as an XPS document using the configured options
            workbook.Save("ConvertedDocument.xps", xpsOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during XPS conversion: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ConvertWorkbookToXps.Run();
    }
}
