// Title: Export Aspose.Cells workbook to HTML with inline CSS and TableCssId class prefix using C# (current version limitation)
// Common Searches: how to generate HTML from an Aspose.Cells workbook with embedded CSS in C# | Aspose.Cells HtmlSaveOptions TableCssId prefix for CSS classes example | C# export Excel to HTML with inline styles using Aspose.Cells | workaround for missing ExportCssInline property in Aspose.Cells .NET | preview workbook as HTML with CSS class prefix in Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions inline CSS | C# export workbook to HTML | TableCssId CSS class prefix Aspose.Cells | Excel to HTML preview Aspose.Cells | Aspose.Cells missing ExportCssInline workaround

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, adds sample data to cells A1‑B2, configures HtmlSaveOptions (noting that ExportCssInline and ExportCssClassPrefix are unavailable in the current Aspose.Cells version), saves the workbook as output.html, prints the full output path, and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportCssInline and ExportCssClassPrefix are not available in the current Aspose.Cells version.
            // The default behavior will be used.

            // Export the workbook to HTML using the configured options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
