// Title: Convert an HTML file to XLSX and transfer <meta> tag values into custom document properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML document, extracts all <meta name='...'> tags with a regular expression, adds each as a custom document property in an Aspose.Cells Workbook, and saves the workbook as an .xlsx file. | Update the sample program to also capture <meta http-equiv='...'> tags and map them to the appropriate built‑in workbook properties (e.g., Author, Title) before exporting to Excel. | Write a reusable method `AddMetaProperties(string html, Workbook workbook)` that parses meta tags from an HTML string and populates the workbook's custom document properties, handling duplicate names gracefully.
// Common Searches: how to import html meta tags as custom properties in an Excel file using Aspose.Cells C# | convert html page to xlsx while preserving meta information with Aspose.Cells | c# regex extract meta name and content from html for Excel export | aspocells load html and set custom document properties from meta tags
// Tags: load html into Aspose.Cells workbook | parse html meta tags using regular expression | add custom document properties to workbook | save workbook as xlsx with custom properties | convert html meta information to Excel properties

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// // Loads an HTML file into an Aspose.Cells Workbook, extracts <meta name='...'> tags via a case‑insensitive regex, adds each as a custom document property (creating or updating as needed), and saves the workbook as an .xlsx file.
class HtmlToExcelWithProperties
{
    static void Main()
    {
        // Paths for input HTML and output Excel files
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        try
        {
            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at path '{htmlPath}'.");
                return;
            }

            // Load the HTML file into a Workbook using Aspose.Cells
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Read the HTML content to extract <meta> tags
            string htmlContent = File.ReadAllText(htmlPath);

            // Regex to match <meta name="..." content="..."> (case‑insensitive)
            string pattern = @"<meta\s+[^>]*name\s*=\s*[""'](?<name>[^""']+)[""'][^>]*content\s*=\s*[""'](?<value>[^""']*)[""'][^>]*>";
            foreach (Match match in Regex.Matches(htmlContent, pattern, RegexOptions.IgnoreCase))
            {
                string propName = match.Groups["name"].Value.Trim();
                string propValue = match.Groups["value"].Value.Trim();

                if (string.IsNullOrEmpty(propName))
                    continue;

                // Add or update the custom document property
                if (workbook.CustomDocumentProperties.Contains(propName))
                {
                    workbook.CustomDocumentProperties[propName].Value = propValue;
                }
                else
                {
                    workbook.CustomDocumentProperties.Add(propName, propValue);
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
