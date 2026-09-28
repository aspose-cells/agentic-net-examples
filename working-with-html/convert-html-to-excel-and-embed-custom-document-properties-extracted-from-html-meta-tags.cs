// Title: Convert an HTML file to an XLSX workbook and import its <meta> tag values as custom document properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML file with Aspose.Cells, extracts every <meta name="..." content="..."> element, and writes each pair as a custom document property in the created Excel workbook. | Create a reusable C# method that accepts an HTML file path, returns a Workbook whose custom properties are populated from the HTML meta tags, and optionally saves it as XLSX. | Show how to extend the example to also map the HTML <title> element to the workbook's built‑in Title property while preserving other meta‑tag properties.
// Common Searches: Aspose.Cells extract HTML meta tags and store them as custom properties in Excel | C# convert HTML page to XLSX while keeping meta information | How to map <meta name> values to Excel custom document properties using Aspose.Cells | Load HTML into Aspose.Cells workbook and add custom properties from meta elements | Preserve HTML metadata when converting to Excel with Aspose.Cells .NET
// Tags: Aspose.Cells HTML to XLSX conversion with custom properties | C# regex extraction of HTML meta tags | populate workbook custom document properties from HTML | load HTML workbook using LoadOptions Html | update or add custom properties in Aspose.Cells workbook

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The sample loads an HTML file into an Aspose.Cells Workbook, uses a regular expression to locate all <meta name="..." content="..."> tags, and adds each name/value pair as a custom document property (updating existing entries when needed). The workbook is then saved as an XLSX file, preserving the HTML metadata inside the Excel document.
class Program
{
    static void Main()
    {
        try
        {
            // Input HTML file path
            string htmlPath = "input.html";

            // Output Excel file path
            string excelPath = "output.xlsx";

            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            // Load the HTML document into an Aspose.Cells workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Read the HTML content as text
            string htmlContent = File.ReadAllText(htmlPath);

            // Regex to find <meta name="..." content="..."> tags (case‑insensitive)
            string pattern = @"<meta\s+[^>]*name\s*=\s*[""'](?<name>[^""']+)[""'][^>]*content\s*=\s*[""'](?<content>[^""']+)[""'][^>]*>";
            var matches = Regex.Matches(htmlContent, pattern, RegexOptions.IgnoreCase);

            // Process each meta tag found
            foreach (Match match in matches)
            {
                string name = match.Groups["name"].Value.Trim();
                string content = match.Groups["content"].Value.Trim();

                if (!string.IsNullOrEmpty(name))
                {
                    var customProps = workbook.CustomDocumentProperties;
                    if (customProps.Contains(name))
                    {
                        customProps[name].Value = content;
                    }
                    else
                    {
                        customProps.Add(name, content);
                    }
                }
            }

            // Save the workbook as an Excel file
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
