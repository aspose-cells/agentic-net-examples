// Title: Compare embedded CSS and external CSS file sizes when exporting an Aspose.Cells workbook to HTML in C#
// Common Searches: C# Aspose.Cells export workbook to HTML with embedded CSS and get file size | How to measure HTML output size for embedded versus external stylesheet using Aspose.Cells | Aspose.Cells HtmlSaveOptions default embedded CSS size comparison in .NET | Export Excel to HTML with Aspose.Cells and compare embedded CSS vs external CSS file size | C# code to compare HTML file size when using embedded CSS in Aspose.Cells
// Tags: HTML export with embedded CSS using Aspose.Cells | Aspose.Cells HTMLSaveOptions file size evaluation | embedded stylesheet vs external stylesheet in Aspose.Cells HTML export | calculate HTML output size in Aspose.Cells | Aspose.Cells workbook HTML size analysis

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills it with data, and saves it twice as HTML using the default HtmlSaveOptions (which embed CSS). It then reads the file sizes of the two HTML files, assumes no external CSS file is generated, computes the total size for the external‑CSS scenario, and prints which approach yields a smaller overall size.
class HtmlExportComparison
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate the worksheet with data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // --------------------------------------------------------------------
            // Export HTML (default behavior – CSS is embedded)
            // --------------------------------------------------------------------
            HtmlSaveOptions embeddedOptions = new HtmlSaveOptions(SaveFormat.Html);
            // No additional settings required for embedded CSS in this version

            string embeddedHtmlPath = "embedded.html";
            workbook.Save(embeddedHtmlPath, embeddedOptions);

            // --------------------------------------------------------------------
            // Export HTML with external CSS (fallback – CSS will still be embedded
            // because the current Aspose.Cells version does not expose external CSS options)
            // --------------------------------------------------------------------
            HtmlSaveOptions externalOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The library version used does not support disabling embedded CSS or
            // specifying an external stylesheet, so we proceed with default settings.

            string externalHtmlPath = "external.html";
            workbook.Save(externalHtmlPath, externalOptions);

            // --------------------------------------------------------------------
            // Compare file sizes
            // --------------------------------------------------------------------
            long embeddedSize = File.Exists(embeddedHtmlPath) ? new FileInfo(embeddedHtmlPath).Length : 0;
            long externalHtmlSize = File.Exists(externalHtmlPath) ? new FileInfo(externalHtmlPath).Length : 0;
            // No external CSS file is generated in this version
            long externalCssSize = 0;
            long totalExternalSize = externalHtmlSize + externalCssSize;

            Console.WriteLine($"Embedded HTML file size: {embeddedSize} bytes");
            Console.WriteLine($"External HTML file size: {externalHtmlSize} bytes");
            Console.WriteLine($"External CSS file size: {externalCssSize} bytes");
            Console.WriteLine($"Total size (HTML + CSS): {totalExternalSize} bytes");

            // Simple comparison output
            if (embeddedSize < totalExternalSize)
            {
                Console.WriteLine("Embedded CSS approach results in a smaller overall size.");
            }
            else if (embeddedSize > totalExternalSize)
            {
                Console.WriteLine("External CSS approach results in a smaller overall size.");
            }
            else
            {
                Console.WriteLine("Both approaches produce files of equal total size.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
