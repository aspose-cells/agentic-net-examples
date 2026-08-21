// Title: C# – Convert HTML to XLSX and embed meta tag values as custom document properties with Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, extracts <meta> name/property and content attributes using a case‑insensitive regex, adds each unique meta entry as a custom document property, and saves the result as an XLSX workbook.
// Keywords: Aspose.Cells HTML to XLSX C# | extract meta tags regex C# | custom document properties Excel | convert web page to Excel | Aspose.Cells workbook custom properties | load HTML workbook Aspose | save as Xlsx C#
// Common Searches: Aspose.Cells load HTML and save as XLSX | add HTML meta tags as custom properties in Excel using C# | extract meta name and content from HTML with regex | convert web page to Excel preserving metadata | C# code to read meta tags and set workbook custom properties
// Developer Intent: Load an HTML document, convert it to an Excel workbook, and store the HTML meta tag values as custom document properties.
// Use Cases: Create Excel reports from web pages while retaining SEO metadata for downstream analytics. | Migrate legacy HTML reports to XLSX, capturing author, date, and other meta information as custom properties. | Run a batch job that processes multiple HTML files, converts each to XLSX, and embeds its meta tags for document‑management systems.
// AI Prompts: Provide C# code that uses Aspose.Cells to load an HTML file, extract all meta name/property tags with a regular expression, add them as custom document properties, and save as XLSX. | Explain how to handle duplicate meta tag names when adding custom document properties with Aspose.Cells. | Suggest improvements to the regex pattern for robust extraction of meta tags with varying attribute order and whitespace.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Loads an HTML file into an Aspose.Cells Workbook, extracts <meta> name/property and content attributes using a case‑insensitive regex, adds each unique meta entry as a custom document property, and saves the result as an XLSX workbook.
class HtmlToExcelWithCustomProperties
{
    static void Main()
    {
        // Paths for source HTML and destination Excel files
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        try
        {
            // Verify that the HTML source file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at path '{htmlPath}'.");
                return;
            }

            // ---------- Load ----------
            // Load the HTML file into a new workbook instance (Aspose.Cells detects the format automatically)
            Workbook workbook = new Workbook(htmlPath);

            // Read the raw HTML text to extract meta tags
            string htmlContent = File.ReadAllText(htmlPath);

            // Regex pattern to capture meta tags with name/property and content attributes
            Regex metaRegex = new Regex(
                @"<meta\s+[^>]*?(?:name|property)\s*=\s*[""'](?<name>[^""']+)[""']\s+[^>]*?content\s*=\s*[""'](?<content>[^""']*)[""']|content\s*=\s*[""'](?<content2>[^""']*)[""']\s+[^>]*?(?:name|property)\s*=\s*[""'](?<name2>[^""']+)[""'])",
                RegexOptions.IgnoreCase);

            // Iterate over all matches and add them as custom document properties
            foreach (Match match in metaRegex.Matches(htmlContent))
            {
                string name = match.Groups["name"].Success ? match.Groups["name"].Value : match.Groups["name2"].Value;
                string value = match.Groups["content"].Success ? match.Groups["content"].Value : match.Groups["content2"].Value;

                // Add the property only if it does not already exist
                if (workbook.CustomDocumentProperties[name] == null)
                {
                    workbook.CustomDocumentProperties.Add(name, value);
                }
            }

            // ---------- Save ----------
            // Save the workbook as an XLSX file
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
