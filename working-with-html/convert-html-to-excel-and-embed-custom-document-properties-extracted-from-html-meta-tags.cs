using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace HtmlToExcelWithCustomProperties
{
    class Program
    {
        static void Main()
        {
            // Paths for source HTML and destination Excel file
            string htmlFilePath = "input.html";
            string excelFilePath = "output.xlsx";

            // Load the HTML file into a workbook (Aspose.Cells can load HTML directly)
            Workbook workbook = new Workbook(htmlFilePath);

            // Read the HTML content as text
            string htmlContent = File.ReadAllText(htmlFilePath);

            // Regex to capture <meta name="..." content="..."> tags (case‑insensitive)
            Regex metaRegex = new Regex(
                @"<meta\s+[^>]*name\s*=\s*[""'](?<name>[^""']+)[""']\s+[^>]*content\s*=\s*[""'](?<content>[^""']+)[""'][^>]*>",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            // Iterate over all matches and add them as custom document properties
            foreach (Match match in metaRegex.Matches(htmlContent))
            {
                string propName = match.Groups["name"].Value.Trim();
                string propValue = match.Groups["content"].Value.Trim();

                // Avoid adding empty names
                if (!string.IsNullOrEmpty(propName))
                {
                    // If the property already exists, update its value; otherwise, add a new one
                    if (workbook.CustomDocumentProperties.Contains(propName))
                    {
                        workbook.CustomDocumentProperties[propName].Value = propValue;
                    }
                    else
                    {
                        workbook.CustomDocumentProperties.Add(propName, propValue);
                    }
                }
            }

            // Save the workbook as an Excel file (XLSX format)
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML converted to Excel and saved to '{excelFilePath}'.");
        }
    }
}