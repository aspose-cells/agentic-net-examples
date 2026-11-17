// Title: Generate printable HTML from an Excel workbook using Aspose.Cells for .NET with an external @media print stylesheet (C#)
// Common Searches: how to export an Excel worksheet to HTML with a separate printable CSS file using Aspose.Cells in C# | Aspose.Cells C# add external stylesheet link to exported HTML for print media | create @media print rules for HTML output generated from Excel with Aspose.Cells | export only the active worksheet to HTML and embed images as Base64 using Aspose.Cells | programmatically insert a CSS link into Aspose.Cells generated HTML in C#
// Tags: Aspose.Cells HTML export external CSS | C# printable HTML from Excel workbook | HtmlSaveOptions active worksheet only | inject link tag into Aspose.Cells HTML output | @media print styling for Aspose.Cells generated HTML

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an Excel file, writes a CSS file containing general table styles and @media print rules, configures HtmlSaveOptions to export only the active worksheet with images as Base64, saves the workbook as HTML, and then injects a <link> element referencing the external stylesheet into the generated HTML's <head> section.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.html";
            const string cssFileName = "printStyles.css";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputFile);

            // CSS content with @media print rules
            string cssContent = @"
                /* General table styling */
                table { border-collapse: collapse; }
                td, th { border: 1px solid #ccc; padding: 5px; }

                /* Print‑specific styling */
                @media print {
                    .no-print { display: none !important; }
                    @page { margin: 1cm; }
                    table { width: 100% !important; font-size: 10pt; }
                }
            ";

            // Write the CSS to an external file
            File.WriteAllText(cssFileName, cssContent);

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true,
                ExportImagesAsBase64 = true
            };

            // Export the workbook to HTML
            workbook.Save(outputFile, htmlOptions);

            // Insert a link to the external CSS file into the generated HTML
            if (File.Exists(outputFile))
            {
                string html = File.ReadAllText(outputFile);
                string linkTag = $"<link rel=\"stylesheet\" type=\"text/css\" href=\"{cssFileName}\" />";
                int headPos = html.IndexOf("<head>", StringComparison.OrdinalIgnoreCase);
                if (headPos != -1)
                {
                    int insertPos = headPos + "<head>".Length;
                    html = html.Insert(insertPos, Environment.NewLine + linkTag + Environment.NewLine);
                    File.WriteAllText(outputFile, html);
                }
            }

            Console.WriteLine($"HTML file '{outputFile}' generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
