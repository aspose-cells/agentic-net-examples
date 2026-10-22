// Title: Export an Aspose.Cells workbook to HTML in C# and confirm each <td> element includes a data‑cell attribute
// AI Prompts: Write C# code that creates a Workbook, saves it as HTML using Aspose.Cells, reads the HTML string, and uses an HTML parser to ensure every <td> tag contains a data‑cell attribute. | Show how to customize the Aspose.Cells HTML export so that a data‑cell attribute is automatically added to each table cell, then validate the attribute with C#.
// Common Searches: how to add custom data-cell attribute to HTML cells when exporting Excel with Aspose.Cells C# | C# verify that all <td> tags have a data-cell attribute after saving workbook as HTML | Aspose.Cells HTML export validation of table cell attributes using regex or HtmlAgilityPack | read generated HTML from Aspose.Cells workbook via MemoryStream in C# | check missing data-cell attributes in Aspose.Cells HTML output
// Tags: Aspose.Cells HTML export custom data-cell attribute | C# validate td attributes with HtmlAgilityPack | regex check td data-cell attribute in generated HTML | MemoryStream read HTML from Aspose.Cells workbook | add custom attribute during Excel to HTML conversion

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example creates a workbook, populates it with sample data, saves it to HTML via a MemoryStream, reads the HTML content, and then scans all <td> elements to verify each contains a data‑cell attribute, reporting success or listing any missing attributes.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a workbook and add some data.
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Row1Col1");
            sheet.Cells["B2"].PutValue("Row1Col2");
            sheet.Cells["A3"].PutValue("Row2Col1");
            sheet.Cells["B3"].PutValue("Row2Col2");

            // 2. Export the workbook to HTML and capture the HTML as a string.
            string htmlContent;
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Html);
                ms.Position = 0;
                using (var reader = new StreamReader(ms))
                {
                    htmlContent = reader.ReadToEnd();
                }
            }

            // 3. Use a regular expression to find all <td> elements.
            var tdMatches = Regex.Matches(htmlContent, @"<td\b[^>]*>", RegexOptions.IgnoreCase);
            if (tdMatches.Count == 0)
            {
                Console.WriteLine("No <td> elements found in the generated HTML.");
                return;
            }

            // 4. Verify each <td> has a 'data-cell' attribute.
            bool allHaveAttribute = true;
            foreach (Match match in tdMatches)
            {
                string tdTag = match.Value;
                // Check for data-cell attribute (case‑insensitive)
                if (!Regex.IsMatch(tdTag, @"\bdata-cell\s*=", RegexOptions.IgnoreCase))
                {
                    allHaveAttribute = false;
                    // Extract inner text roughly (optional, not required for attribute check)
                    Console.WriteLine($"Missing data-cell attribute in tag: {tdTag}");
                }
            }

            if (allHaveAttribute)
            {
                Console.WriteLine("All table cells contain the 'data-cell' attribute.");
            }
            else
            {
                Console.WriteLine("Some table cells are missing the 'data-cell' attribute.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
