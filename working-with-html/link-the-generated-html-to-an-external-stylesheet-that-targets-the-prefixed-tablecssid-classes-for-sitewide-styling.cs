// Title: C# – Link an external CSS file to Aspose.Cells HTML using TableCssId prefix
// Description: Creates a workbook, sets HtmlSaveOptions.TableCssId to a custom prefix, exports worksheet CSS to a separate .css file, and programmatically injects a <link> tag into the generated HTML so the site‑wide stylesheet styles .{prefix}-table, .{prefix}-tr, and .{prefix}-td elements.
// Keywords: Aspose.Cells HTML export C# | TableCssId prefix | external CSS link | ExportWorksheetCSSSeparately | inject link tag C# | Excel to HTML styling
// Common Searches: how to add external stylesheet to Aspose.Cells HTML output | use TableCssId to style generated HTML tables | programmatically insert <link> into saved HTML C# | export worksheet CSS separately Aspose.Cells
// Developer Intent: Add a shared stylesheet that targets the prefixed CSS classes produced by Aspose.Cells when saving a workbook as HTML.
// Use Cases: Generate HTML reports from Excel files with consistent site‑wide table styling. | Reuse a single CSS file across multiple Excel‑to‑HTML conversions by applying a common class prefix. | Automate post‑processing of saved HTML to embed a <link> element, removing manual edits.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as HTML, creates a CSS file for .site-table, .site-tr, .site-td, and inserts a <link> tag into the HTML head. | Explain how HtmlSaveOptions.TableCssId prefixes generated CSS classes and how to reference an external stylesheet for global styling. | Show a C# snippet that reads a saved HTML file, finds the </head> tag, and injects a <link rel="stylesheet" href="..."> element.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, sets HtmlSaveOptions.TableCssId to a custom prefix, exports worksheet CSS to a separate .css file, and programmatically injects a <link> tag into the generated HTML so the site‑wide stylesheet styles .{prefix}-table, .{prefix}-tr, and .{prefix}-td elements.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Prefix for generated CSS classes (e.g., .site-tr, .site-td)
        saveOptions.TableCssId = "site";
        // Export worksheet CSS to a separate file instead of embedding it
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Define output file names
        string htmlPath = "output.html";
        string externalCssPath = "site-styles.css";

        // Save the workbook as HTML (a separate CSS file will also be generated)
        workbook.Save(htmlPath, saveOptions);

        // Create a site‑wide stylesheet that targets the prefixed classes
        string cssContent = @"
.site-table { border-collapse: collapse; width: 100%; }
.site-tr { background-color: #f9f9f9; }
.site-td { padding: 8px; border: 1px solid #ccc; }
";
        File.WriteAllText(externalCssPath, cssContent);

        // Insert a <link> tag into the generated HTML to reference the external stylesheet
        string html = File.ReadAllText(htmlPath);
        int headCloseIdx = html.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
        if (headCloseIdx >= 0)
        {
            string linkTag = $@"<link rel=""stylesheet"" type=""text/css"" href=""{externalCssPath}"">";
            html = html.Insert(headCloseIdx, linkTag + Environment.NewLine);
            File.WriteAllText(htmlPath, html);
        }

        Console.WriteLine("HTML file saved and linked to external stylesheet.");
    }
}
