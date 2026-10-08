// Title: Convert an Excel workbook containing WordArt to HTML and minify the embedded CSS gradient fills with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with WordArt, saves it as HTML using Aspose.Cells, extracts the <style> element, and compresses the CSS gradient definitions. | Demonstrate how to replace the original style block in Aspose.Cells‑generated HTML with a whitespace‑optimized version while keeping the document structure intact.
// Common Searches: how to export WordArt from Excel to HTML using Aspose.Cells .NET | minify CSS inside the style tag of Aspose.Cells generated HTML | remove comments and extra spaces from Aspose.Cells HTML CSS output | compress gradient fill CSS when converting Excel to HTML with Aspose.Cells | C# example for extracting and rewriting style block in Aspose.Cells HTML save
// Tags: Aspose.Cells WordArt HTML conversion | C# CSS minification Aspose.Cells output | gradient fill CSS compression .NET | modify style element Aspose.Cells HTML | HtmlSaveOptions CSS optimization

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program loads an Excel workbook that contains WordArt, saves it to HTML using Aspose.Cells, extracts the generated <style> block, minifies the CSS by stripping comments and excess whitespace, reinserts the compressed style tag, and writes the final HTML to a file.
class WordArtToHtmlConverter
{
    static void Main()
    {
        // Load the workbook that contains WordArt.
        // Replace "input.xlsx" with the path to your source workbook.
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Export WordArt as part of the HTML (default behavior).
        // You can adjust options here if needed, e.g., htmlOptions.ExportChartAsImage = false;

        // Save the workbook to a memory stream to capture the generated HTML.
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);
            // Convert the stream content to a UTF‑8 string.
            string htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());

            // -----------------------------------------------------------------
            // Minify the CSS that defines gradient fills.
            // Aspose.Cells embeds CSS inside a <style> tag. We'll extract that
            // block, minify it (remove whitespace, line‑breaks, comments), and
            // replace the original block with the minified version.
            // -----------------------------------------------------------------

            // Regex to capture the first <style>...</style> block (single‑line mode).
            const string stylePattern = @"<style[^>]*>(.*?)</style>";
            Match styleMatch = Regex.Match(htmlContent, stylePattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

            if (styleMatch.Success)
            {
                string originalCss = styleMatch.Groups[1].Value;

                // Simple CSS minification:
                // 1. Remove CSS comments.
                // 2. Remove line breaks, tabs and excess spaces.
                // 3. Collapse multiple spaces into a single space.
                string minifiedCss = Regex.Replace(originalCss, @"/\*.*?\*/", string.Empty, RegexOptions.Singleline);
                minifiedCss = Regex.Replace(minifiedCss, @"\s+", " ");          // collapse whitespace
                minifiedCss = minifiedCss.Replace(" }", "}").Replace("{ ", "{"); // tidy braces
                minifiedCss = minifiedCss.Trim();

                // Re‑inject the minified CSS back into the HTML.
                string minifiedStyleTag = $"<style>{minifiedCss}</style>";
                htmlContent = Regex.Replace(htmlContent, stylePattern, minifiedStyleTag, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            }

            // Save the final HTML with minified CSS to a file.
            // Replace "output.html" with your desired output path.
            File.WriteAllText("output.html", htmlContent, Encoding.UTF8);
        }

        Console.WriteLine("Conversion completed. HTML saved to output.html");
    }
}
