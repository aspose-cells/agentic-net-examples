// Title: C# – Convert HTML to XLSX and Embed Meta Tag Values as Custom Document Properties with Aspose.Cells
// Description: Loads an HTML file directly into an Aspose.Cells Workbook, parses <meta name="..." content="..."> tags using a case‑insensitive regex, adds each name/value pair as a string custom document property (replacing duplicates), and saves the result as an XLSX workbook.
// Keywords: Aspose.Cells | C# | HTML to Excel conversion | custom document properties | meta tag extraction | regex meta parsing | Workbook.Save Xlsx | batch HTML to XLSX | SEO metadata in Excel
// Common Searches: How to read HTML meta tags and store them in Excel with Aspose.Cells | C# convert HTML file to XLSX and keep meta information | Aspose.Cells add custom document properties from <meta> tags | Extract meta data from web page and embed in Excel workbook | Bulk convert HTML to Excel preserving SEO metadata
// Developer Intent: Generate an Excel workbook from an HTML document and automatically embed the page's meta tag values as custom document properties.
// Use Cases: Create audit‑ready Excel reports that retain author, title, or version information from web pages. | Automate batch conversion of dozens of HTML files while capturing SEO meta data for compliance tracking. | Build a data‑import pipeline that reads HTML templates, extracts configuration values from meta tags, and stores them as workbook properties for downstream processing.
// AI Prompts: Write C# code using Aspose.Cells and HtmlAgilityPack to load an HTML file, extract all <meta> tags without regex, add them as custom document properties, and save as XLSX. | Explain how to update existing custom document properties when the same meta name appears in multiple HTML files during bulk conversion. | Provide a step‑by‑step guide with error handling for converting HTML to Excel with Aspose.Cells and embedding extracted meta information as string properties.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace HtmlToExcelWithCustomProperties
{
    // Loads an HTML file directly into an Aspose.Cells Workbook, parses <meta name="..." content="..."> tags using a case‑insensitive regex, adds each name/value pair as a string custom document property (replacing duplicates), and saves the result as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlFilePath = "input.html";

            // Load the HTML file into a workbook (Aspose.Cells can load HTML directly)
            Workbook workbook = new Workbook(htmlFilePath);

            // Read the HTML content as text
            string htmlContent = File.ReadAllText(htmlFilePath);

            // Regular expression to match <meta name="..." content="..."> tags (case‑insensitive)
            string metaPattern = @"<meta\s+[^>]*name\s*=\s*[""'](?<name>[^""']+)[""']\s+[^>]*content\s*=\s*[""'](?<content>[^""']*)[""'][^>]*>";
            foreach (Match match in Regex.Matches(htmlContent, metaPattern, RegexOptions.IgnoreCase))
            {
                string propName = match.Groups["name"].Value.Trim();
                string propValue = match.Groups["content"].Value.Trim();

                if (!string.IsNullOrEmpty(propName))
                {
                    // Add or update a custom document property with the extracted name/value
                    // If the property already exists, remove it first to avoid duplication
                    if (workbook.CustomDocumentProperties.Contains(propName))
                    {
                        workbook.CustomDocumentProperties.Remove(propName);
                    }

                    // All values are stored as strings; you can change the type if needed
                    workbook.CustomDocumentProperties.Add(propName, propValue);
                }
            }

            // Save the workbook as an Excel file (XLSX format)
            string excelOutputPath = "output.xlsx";
            workbook.Save(excelOutputPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML converted to Excel and saved to '{excelOutputPath}'.");
            Console.WriteLine("Custom document properties extracted from meta tags have been embedded.");
        }
    }
}
