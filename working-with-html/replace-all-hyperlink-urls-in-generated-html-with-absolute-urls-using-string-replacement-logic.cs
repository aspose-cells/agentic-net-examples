// Title: Replace relative hyperlink URLs with absolute URLs in Aspose.Cells‑generated HTML using C# and Regex
// AI Prompts: Generate C# code that reads an HTML file produced by Aspose.Cells, finds href attributes lacking a scheme, and rewrites them to full URLs using a supplied base address. | Show a C# Regex pattern and replacement logic to transform all relative links in Aspose.Cells HTML output into absolute links.
// Common Searches: C# how to convert relative hrefs to absolute URLs in HTML saved by Aspose.Cells | regex replace href without http in Aspose.Cells generated HTML | post‑process Aspose.Cells HTML to update hyperlink paths programmatically | add base URL to relative links after saving workbook as HTML in .NET | Aspose.Cells HTML output modify hyperlink URLs with C#
// Tags: Aspose.Cells HTML hyperlink absolute conversion | C# regex replace relative href | post‑process Aspose.Cells HTML links | generate absolute URLs from base path C# | modify workbook HTML hyperlinks programmatically

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example creates a workbook, adds a relative hyperlink, saves it as HTML, reads the HTML file, uses a regular expression to locate href attributes that do not start with http/https, combines each relative URL with a specified base URL to form an absolute URL, writes the updated HTML to a new file, and outputs the path of the modified file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        sheet.Cells["A1"].PutValue("Link to page");
        // Add a relative hyperlink (e.g., "page.html")
        sheet.Hyperlinks.Add(0, 0, 1, 1, "page.html");

        // Save the workbook as HTML
        string htmlPath = "output.html";
        workbook.Save(htmlPath, SaveFormat.Html);

        // Read the generated HTML
        string htmlContent = File.ReadAllText(htmlPath);

        // Define the base URL to convert relative links to absolute ones
        string baseUrl = "https://example.com/";

        // Replace href attributes that contain relative URLs with absolute URLs
        // This regex finds href="..." where the URL does NOT start with http:// or https://
        string pattern = @"href\s*=\s*""(?!(?:http|https)://)([^""]*)""";
        string replacedHtml = Regex.Replace(htmlContent, pattern, m =>
        {
            string relativeUrl = m.Groups[1].Value;
            // Combine base URL with the relative part
            string absoluteUrl = new Uri(new Uri(baseUrl), relativeUrl).ToString();
            return $"href=\"{absoluteUrl}\"";
        }, RegexOptions.IgnoreCase);

        // Write the modified HTML back to file (or to a new file)
        string modifiedHtmlPath = "output_absolute.html";
        File.WriteAllText(modifiedHtmlPath, replacedHtml);

        Console.WriteLine($"HTML with absolute URLs saved to: {modifiedHtmlPath}");
    }
}
