// Title: C# – Convert Aspose.Cells HTML hyperlinks from relative to absolute URLs
// Description: Creates a workbook, adds relative and absolute hyperlinks, saves it as HTML, reads the output, detects non‑absolute links, builds full URLs using a base address, replaces the href attributes via string replacement, and writes a new HTML file with all hyperlinks fully qualified.
// Keywords: Aspose.Cells C# hyperlink absolute URL | convert relative link to absolute in HTML export | HtmlSaveOptions replace href | C# string replace href Aspose.Cells | post‑process Aspose.Cells HTML links
// Common Searches: how to make Aspose.Cells exported HTML links absolute | replace relative hyperlinks in Aspose.Cells HTML output | C# convert Excel hyperlink to full URL after saving as HTML | Aspose.Cells post‑processing href attributes | C# generate absolute URLs from workbook hyperlinks
// Developer Intent: Transform relative hyperlink addresses in an Aspose.Cells‑generated HTML file into absolute URLs using a specified base URL.
// Use Cases: Publish an Excel‑derived HTML report where every link must resolve from a public domain. | Create downloadable HTML versions of spreadsheets that reference resources hosted on a known website. | Automate post‑export processing to ensure all worksheet hyperlinks are fully qualified before distribution.
// AI Prompts: Generate C# code that iterates over Worksheet.Hyperlinks, detects relative addresses, builds absolute URLs with a base path, and replaces the corresponding href values in the saved HTML string. | Provide a reusable method that accepts an Aspose.Cells HTML file path and a base URL, returning the HTML content with all relative hyperlinks converted to absolute URLs. | Explain how to check whether a hyperlink address is already absolute before performing string replacement in the exported HTML.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHyperlinkAbsoluteUrlDemo
{
    // Creates a workbook, adds relative and absolute hyperlinks, saves it as HTML, reads the output, detects non‑absolute links, builds full URLs using a base address, replaces the href attributes via string replacement, and writes a new HTML file with all hyperlinks fully qualified.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Relative Link");
            sheet.Cells["A2"].PutValue("Absolute Link");

            // Add a relative hyperlink (e.g., a local HTML page)
            sheet.Hyperlinks.Add("A1", 1, 1, "pages/page2.html");

            // Add an absolute hyperlink (already full URL)
            sheet.Hyperlinks.Add("A2", 1, 1, "https://www.google.com");

            // Save the workbook as HTML (default generates relative links)
            string htmlPath = "output.html";
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            workbook.Save(htmlPath, saveOptions);

            // Load the generated HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Define the base URL to prepend to relative hyperlinks
            string baseUrl = "https://www.example.com/";

            // Iterate through all hyperlinks in the worksheet
            foreach (Hyperlink link in sheet.Hyperlinks)
            {
                string originalAddress = link.Address;

                // Determine if the address is already absolute
                bool isAbsolute = originalAddress.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                                  originalAddress.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

                if (!isAbsolute)
                {
                    // Build the absolute URL
                    string absoluteAddress = new Uri(new Uri(baseUrl), originalAddress).ToString();

                    // Replace the href attribute in the HTML content
                    // The HTML exporter uses the address as the value of href attribute
                    string hrefPattern = $"href=\"{originalAddress}\"";
                    string hrefReplacement = $"href=\"{absoluteAddress}\"";
                    htmlContent = htmlContent.Replace(hrefPattern, hrefReplacement);
                }
            }

            // Save the modified HTML with absolute URLs
            string modifiedHtmlPath = "output_absolute.html";
            File.WriteAllText(modifiedHtmlPath, htmlContent);

            Console.WriteLine($"Original HTML saved to: {htmlPath}");
            Console.WriteLine($"Modified HTML with absolute URLs saved to: {modifiedHtmlPath}");
        }
    }
}
