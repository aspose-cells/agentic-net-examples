// Title: Create an HTML image map from Excel worksheets and replace cell click links with custom URLs using Aspose.Cells for .NET
// AI Prompts: Generate an HTML file that contains an image map for each worksheet in an Excel workbook with Aspose.Cells, then edit the HTML to swap the default CellClick JavaScript links for selected cells with your own external URLs. | Load a .xlsx workbook, call Workbook.Save using SaveFormat.Html to produce an image map, read the resulting HTML, locate href attributes that invoke javascript:CellClick for cells such as A1 or B2, and replace them with absolute hyperlinks to the desired resources.
// Common Searches: how to export Excel worksheets to HTML with clickable image map using Aspose.Cells .NET | replace Aspose.Cells generated CellClick JavaScript links with custom URLs | customize HTML image map generated from workbook pages in C# | Aspose.Cells save workbook as HTML and edit cell hyperlinks programmatically | map specific Excel cells to external web pages in Aspose.Cells HTML output
// Tags: Aspose.Cells HTML image map generation | custom cell hyperlink replacement Aspose.Cells | save workbook as HTML with image map .NET | modify generated HTML from Aspose.Cells | map Excel cell to external URL C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, saves it as HTML (which automatically creates an image map for each sheet), reads the generated HTML file, replaces the default CellClick JavaScript links for specified cells (e.g., A1, B2) with custom external URLs, and writes the updated HTML back to disk.
class HtmlImageMapGenerator
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string htmlPath = "WorkbookWithMap.html";

            // Verify source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source workbook not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as HTML (default options generate an image map and images)
            workbook.Save(htmlPath, SaveFormat.Html);

            // OPTIONAL: Map specific cell addresses to custom resource URLs
            var cellToUrl = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "A1", "https://example.com/resource1" },
                { "B2", "https://example.com/resource2" }
                // Add more mappings as needed
            };

            // Verify the HTML file was created before attempting to modify it
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Generated HTML not found: {htmlPath}");
                return;
            }

            // Read the generated HTML
            string htmlContent;
            try
            {
                htmlContent = File.ReadAllText(htmlPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read HTML file: {ex.Message}");
                return;
            }

            // Replace default cell links with custom URLs
            foreach (var kvp in cellToUrl)
            {
                string cellAddress = kvp.Key;
                string url = kvp.Value;

                // Pattern matches: href="javascript:CellClick('Sheet1!A1')"
                string pattern = $@"href\s*=\s*""javascript:CellClick\('([^']+)!{Regex.Escape(cellAddress)}'\)""";
                htmlContent = Regex.Replace(htmlContent, pattern, $"href=\"{url}\"", RegexOptions.IgnoreCase);
            }

            // Save the modified HTML back to file
            try
            {
                File.WriteAllText(htmlPath, htmlContent);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write modified HTML: {ex.Message}");
                return;
            }

            Console.WriteLine("HTML image map generated and customized successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
