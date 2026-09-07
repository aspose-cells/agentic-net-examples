// Title: Export an Excel workbook to HTML with a custom CSS file that overrides cell colors using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, creates a custom.css defining cell background and text colors, saves the workbook as HTML, and inserts a <link> tag for the stylesheet into the generated HTML. | Demonstrate how to use HtmlSaveOptions to export a workbook to HTML and then programmatically modify the output file to reference an external CSS file that forces specific cell styling.
// Common Searches: aspnet export excel to html with custom stylesheet using Aspose.Cells | how to change cell background colors in HTML output from Aspose.Cells | add external CSS link to Aspose.Cells generated HTML file C# | override default HTML cell styles when saving workbook as HTML with Aspose.Cells | save workbook as html and apply custom css with Aspose.Cells .NET
// Tags: Aspose.Cells HTML export with custom stylesheet | HtmlSaveOptions external CSS integration | override default cell colors in HTML output | add stylesheet reference to generated HTML C# | Excel to HTML conversion using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input.xlsx file, loads it with Aspose.Cells, writes a custom.css that forces a light gray background and black text for table cells, saves the workbook as output.html using HtmlSaveOptions, reads the generated HTML, injects a <link> tag pointing to custom.css inside the <head> (or prepends it if <head> is missing), and writes the updated HTML back to disk.
class ExportExcelToHtmlWithCustomCss
{
    static void Main()
    {
        try
        {
            // Verify that the source Excel file exists
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Define and write the custom CSS file
            const string cssFilePath = "custom.css";
            string cssContent = @"
table { border-collapse: collapse; }
td {
    background-color: #f0f0f0 !important; /* Override cell background */
    color: #000000 !important;           /* Override cell text color */
    border: 1px solid #cccccc;
    padding: 5px;
}
";
            File.WriteAllText(cssFilePath, cssContent);

            // Save the workbook as HTML using default options
            const string htmlPath = "output.html";
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            workbook.Save(htmlPath, saveOptions);

            // Inject a link to the external CSS file into the generated HTML
            if (File.Exists(htmlPath))
            {
                string html = File.ReadAllText(htmlPath);
                string linkTag = $"<link rel=\"stylesheet\" type=\"text/css\" href=\"{cssFilePath}\" />";
                int headPos = html.IndexOf("<head>", StringComparison.OrdinalIgnoreCase);
                if (headPos >= 0)
                {
                    int insertPos = headPos + "<head>".Length;
                    html = html.Insert(insertPos, Environment.NewLine + linkTag);
                }
                else
                {
                    // Fallback: prepend the link tag if <head> is missing
                    html = linkTag + Environment.NewLine + html;
                }
                File.WriteAllText(htmlPath, html);
            }

            Console.WriteLine("Export completed. HTML file and custom CSS have been generated.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
