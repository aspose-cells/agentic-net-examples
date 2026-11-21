// Title: How to prefix a custom TableCssId to self‑closing HTML tags generated from an Excel workbook using Aspose.Cells in C#
// Common Searches: C# add custom id prefix to img, br, input tags in Aspose.Cells HTML export | regex replace self‑closing tags with prefixed id attribute after saving workbook as HTML | Aspose.Cells save Excel to HTML and inject TableCssId into element IDs using C# | modify generated HTML from Aspose.Cells to include TableCssId on self‑closing elements | add id attribute to self‑closing tags in HTML output of Aspose.Cells workbook conversion
// Tags: Aspose.Cells HTML export id prefix | C# regex modify self‑closing tags | add id attribute to img br input elements | custom TableCssId injection in generated HTML | Excel to HTML conversion with element identifier

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, saves it to HTML, then uses a regular expression to find self‑closing tags (e.g., <img />, <br />, <input />). It prefixes any existing id attribute with the constant TableCssId or adds a new id attribute based on the tag name, reconstructs the tag, and writes the processed HTML to an output file.
class HtmlSelfClosingTagProcessor
{
    static void Main()
    {
        const string TableCssId = "MyTable";
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.html";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook to HTML in a memory stream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, SaveFormat.Html);
                string html = Encoding.UTF8.GetString(htmlStream.ToArray());

                // Regex to find self‑closing tags (e.g., <img />, <br />, <input />, etc.)
                // Captures: 1 = tag name, 2 = existing attributes (if any)
                string pattern = @"<(?<tag>\w+)(?<attrs>[^>]*?)\s*/>";

                // Replace each self‑closing tag with the same tag plus a prefixed identifier
                string processedHtml = Regex.Replace(
                    html,
                    pattern,
                    match =>
                    {
                        string tagName = match.Groups["tag"].Value;
                        string attrs = match.Groups["attrs"].Value;

                        // If an id attribute exists, prefix its value
                        if (Regex.IsMatch(attrs, @"\bid\s*=", RegexOptions.IgnoreCase))
                        {
                            attrs = Regex.Replace(
                                attrs,
                                @"\bid\s*=\s*[""'](?<idValue>[^""']*)[""']",
                                m => $"id=\"{TableCssId}_{m.Groups["idValue"].Value}\"",
                                RegexOptions.IgnoreCase);
                        }
                        else
                        {
                            // Add a new id attribute
                            attrs = $"{attrs} id=\"{TableCssId}_{tagName}\"";
                        }

                        // Reconstruct the self‑closing tag
                        return $"<{tagName}{attrs} />";
                    },
                    RegexOptions.IgnoreCase);

                // Write the processed HTML to a file
                File.WriteAllText(outputPath, processedHtml, Encoding.UTF8);
                Console.WriteLine($"HTML saved to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
