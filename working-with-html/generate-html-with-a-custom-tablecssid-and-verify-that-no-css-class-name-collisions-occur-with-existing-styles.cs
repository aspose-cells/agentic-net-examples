// Title: Generate HTML from an Aspose.Cells workbook with a custom TableCssId and automatically avoid CSS class name collisions in C#
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as HTML, sets HtmlSaveOptions.TableCssId to a given value, reads the generated file, extracts the CSS block, and if the TableCssId appears as a CSS class selector, modifies the ID and repeats until it is unique. | Create a .NET helper method that parses the first <style> element from an HTML string, returns all CSS class names, and validates a custom TableCssId against those names to prevent collisions.
// Common Searches: Aspose.Cells C# export workbook to HTML with custom TableCssId and avoid ID conflicts | how to check for CSS class name collision with TableCssId in Aspose.Cells HTML output | C# loop to generate unique TableCssId when exporting Excel to HTML using Aspose.Cells | extract CSS from generated HTML file using regex in .NET for Aspose.Cells export | prevent duplicate element IDs in HTML saved by Aspose.Cells HtmlSaveOptions
// Tags: Aspose.Cells HtmlSaveOptions TableCssId custom | C# detect CSS class collision in generated HTML | unique TableCssId generation Aspose.Cells | extract embedded CSS from HTML .NET | retry mechanism for HTML export Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, exports it to HTML with Aspose.Cells while assigning a custom TableCssId, reads the resulting HTML to extract the embedded CSS, checks whether the TableCssId conflicts with any CSS class selector, and if a collision is found appends a numeric suffix and retries up to five times, finally reporting the unique TableCssId or an error.
class HtmlExportWithCustomTableId
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Desired custom Table CSS Id
            const string baseTableId = "myCustomTable";
            string tableId = baseTableId;

            // Path settings
            string outputHtmlPath = "output.html";

            // Maximum attempts to avoid infinite loop
            const int maxAttempts = 5;
            int attempt = 0;
            bool collisionDetected;

            do
            {
                // Configure HTML save options (only properties supported by the current Aspose.Cells version)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = true,
                    TableCssId = tableId // set custom Table CSS Id
                };

                // Save workbook as HTML (CSS will be embedded in the HTML file)
                workbook.Save(outputHtmlPath, saveOptions);

                // Ensure the HTML file exists before reading
                if (!File.Exists(outputHtmlPath))
                {
                    Console.WriteLine("Failed to generate the HTML file.");
                    return;
                }

                // Load generated HTML content
                string htmlContent = File.ReadAllText(outputHtmlPath);

                // Extract CSS block from <style> tag
                string cssContent = ExtractCssFromHtml(htmlContent);

                // Find all CSS class selectors in the extracted CSS
                var classMatches = Regex.Matches(cssContent, @"\.([A-Za-z0-9_-]+)");
                collisionDetected = false;
                foreach (Match match in classMatches)
                {
                    string className = match.Groups[1].Value;
                    if (string.Equals(className, tableId, StringComparison.OrdinalIgnoreCase))
                    {
                        // Collision found
                        collisionDetected = true;
                        break;
                    }
                }

                if (collisionDetected)
                {
                    // Modify the TableCssId to avoid collision and retry
                    attempt++;
                    tableId = $"{baseTableId}_{attempt}";
                }

            } while (collisionDetected && attempt < maxAttempts);

            if (collisionDetected)
            {
                Console.WriteLine("Unable to generate a unique TableCssId after multiple attempts.");
            }
            else
            {
                Console.WriteLine($"HTML exported successfully with TableCssId = \"{tableId}\"");
                Console.WriteLine($"HTML file: {Path.GetFullPath(outputHtmlPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to extract CSS content from the first <style>...</style> block in the HTML
    private static string ExtractCssFromHtml(string html)
    {
        var styleMatch = Regex.Match(html, @"<style[^>]*>(.*?)</style>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        return styleMatch.Success ? styleMatch.Groups[1].Value : string.Empty;
    }
}
