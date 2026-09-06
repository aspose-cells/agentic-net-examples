// Title: Verify that each Excel worksheet name is rendered as an <h1>‑<h6> heading when saving to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook using Aspose.Cells, saves it as HTML, and checks that every worksheet name appears inside an <h1>‑<h6> tag. | Create a C# method that parses the generated HTML, extracts all heading texts with a regular expression, and returns any worksheet titles that are missing from the headings.
// Common Searches: aspocells verify worksheet titles in generated html headings | c# extract h1 h2 tags from aspocells html output | ensure Excel sheet names become headings in saved html using aspocells | regex to find heading tags in aspocells html file c#
// Tags: Aspose.Cells HTML heading verification | worksheet title to heading mapping | C# regex h1‑h6 extraction | validate Excel sheet titles in HTML | save workbook as HTML Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program loads an Excel workbook, saves it as HTML with Aspose.Cells, reads the HTML file, extracts all <h1>‑<h6> tags using a regular expression, and verifies that each worksheet name is present among those headings, reporting pass or fail for each sheet.
class HtmlHeadingVerifier
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Save the workbook as HTML
        string htmlPath = "output.html";
        workbook.Save(htmlPath, SaveFormat.Html);

        // Read the generated HTML content
        string htmlContent = File.ReadAllText(htmlPath);

        // Prepare a regex to find heading tags (h1 to h6)
        Regex headingRegex = new Regex(@"<h([1-6])\b[^>]*>(.*?)</h\1>", RegexOptions.IgnoreCase);

        // Extract all heading texts from the HTML
        var headings = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match match in headingRegex.Matches(htmlContent))
        {
            // Group 2 contains the inner text of the heading
            string headingText = match.Groups[2].Value.Trim();
            if (!string.IsNullOrEmpty(headingText))
                headings.Add(headingText);
        }

        // Verify each worksheet title appears as a heading in the HTML
        bool allMatch = true;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            string sheetName = sheet.Name;
            if (headings.Contains(sheetName))
            {
                Console.WriteLine($"PASS: Worksheet title \"{sheetName}\" found as a heading.");
            }
            else
            {
                Console.WriteLine($"FAIL: Worksheet title \"{sheetName}\" NOT found as a heading.");
                allMatch = false;
            }
        }

        if (allMatch)
        {
            Console.WriteLine("All worksheet titles are present as heading tags in the generated HTML.");
        }
        else
        {
            Console.WriteLine("Some worksheet titles are missing heading tags in the generated HTML.");
        }
    }
}
