// Title: Convert an HTML file to an Excel workbook and embed HTML meta tag values as custom document properties using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load HTML from string and save as XLSX with custom properties | How to read meta tags from an HTML file and add them to an Excel workbook's custom document properties in .NET | Aspose.Cells HtmlLoadOptions example for extracting meta information and storing it in workbook properties | Convert HTML to Excel while preserving meta data as custom properties using C# | Using regex to parse HTML meta tags and set Workbook.CustomDocumentProperties in Aspose.Cells
// Tags: Aspose.Cells HtmlLoadOptions load HTML string | Workbook.CustomDocumentProperties add meta tag values | C# regex extract HTML meta name content | HTML to XLSX conversion with custom properties | Save workbook as Xlsx using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The sample reads an HTML file, extracts <meta> name/content pairs with a regular expression, loads the HTML into an Aspose.Cells Workbook via HtmlLoadOptions, adds each pair as a custom document property, and saves the result as an .xlsx workbook.
class HtmlToExcelWithCustomProperties
{
    static void Main(string[] args)
    {
        try
        {
            // Input HTML file path
            string htmlPath = @"C:\Input\sample.html";

            // Output Excel file path
            string excelPath = @"C:\Output\sample.xlsx";

            // Verify input file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input HTML file not found: {htmlPath}");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load HTML content
            string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);

            // Parse HTML to extract meta tags (name/content)
            var metaProperties = ExtractMetaProperties(htmlContent);

            // Load HTML into Aspose.Cells workbook
            var loadOptions = new HtmlLoadOptions();
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                Workbook workbook = new Workbook(stream, loadOptions);

                // Add extracted meta tags as custom document properties
                foreach (var kvp in metaProperties)
                {
                    if (workbook.CustomDocumentProperties.Contains(kvp.Key))
                        workbook.CustomDocumentProperties[kvp.Key].Value = kvp.Value;
                    else
                        workbook.CustomDocumentProperties.Add(kvp.Key, kvp.Value);
                }

                // Save the workbook to Excel format
                workbook.Save(excelPath, SaveFormat.Xlsx);
            }

            Console.WriteLine($"Excel file created successfully at: {excelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <param name="html">The raw HTML string.</param>
    /// <returns>A dictionary of meta name => content.</returns>
    private static Dictionary<string, string> ExtractMetaProperties(string html)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Regex to match <meta ... name="..." ... content="..." ...>
        string pattern = @"<meta\s+[^>]*?(?:name\s*=\s*[""'](?<name>[^""']+)[""']|content\s*=\s*[""'](?<content>[^""']+)[""'])(?:\s+[^>]*?(?:name\s*=\s*[""'](?<name2>[^""']+)[""']|content\s*=\s*[""'](?<content2>[^""']+)[""']))?[^>]*?>";
        foreach (Match match in Regex.Matches(html, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline))
        {
            string name = match.Groups["name"].Success ? match.Groups["name"].Value :
                          match.Groups["name2"].Success ? match.Groups["name2"].Value : null;
            string content = match.Groups["content"].Success ? match.Groups["content"].Value :
                             match.Groups["content2"].Success ? match.Groups["content2"].Value : null;

            if (!string.IsNullOrWhiteSpace(name) && content != null)
            {
                name = name.Trim();
                content = content.Trim();
                result[name] = content; // later values overwrite earlier ones
            }
        }

        return result;
    }
}
