// Title: Export an Excel workbook containing WordArt to HTML and minify the generated CSS gradient fills with Aspose.Cells for .NET
// Common Searches: how to export Excel WordArt to HTML using Aspose.Cells C# | minify CSS gradient definitions in HTML produced by Aspose.Cells | save Excel workbook as HTML with embedded base64 images in C# | remove whitespace from style tags in Aspose.Cells HTML output | convert workbook with WordArt to HTML and compress CSS with regex
// Tags: Aspose.Cells WordArt HTML export | C# regex minify CSS in HTML | embed images as base64 in Aspose.Cells HTML | compress gradient fill CSS in exported HTML | MemoryStream handling for Aspose.Cells HTML save

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook that contains WordArt, uses Aspose.Cells HtmlSaveOptions to export all worksheets with images embedded as base64, captures the HTML output via a MemoryStream, then applies a regex routine to collapse whitespace and tighten gradient‑function syntax inside <style> tags, and finally writes the minified HTML to disk.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithWordArt.xlsx";
            const string outputPath = "Output.html";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook that contains WordArt.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false, // Export all worksheets.
                ExportImagesAsBase64 = true         // Embed images directly.
                // WordArt is exported by default; no explicit property required.
                // CSS will be embedded; no separate property needed.
            };

            // Save the workbook to a memory stream to obtain the HTML as a string.
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;
                string html = new StreamReader(ms).ReadToEnd();

                // Minify the CSS that defines gradient fills.
                string minifiedHtml = MinifyCssInHtml(html);

                // Write the final HTML to disk.
                File.WriteAllText(outputPath, minifiedHtml);
            }

            Console.WriteLine($"HTML successfully generated at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Removes unnecessary whitespace from CSS inside <style> tags,
    // focusing on gradient fill definitions.
    static string MinifyCssInHtml(string html)
    {
        return Regex.Replace(html,
            @"<style[^>]*>(.*?)</style>",
            match =>
            {
                string css = match.Groups[1].Value;

                // Collapse all whitespace to a single space.
                css = Regex.Replace(css, @"\s+", " ");

                // Remove spaces around CSS punctuation.
                css = Regex.Replace(css, @"\s*([{}:;,])\s*", "$1");

                // Tighten gradient function syntax.
                css = Regex.Replace(css, @"linear-gradient\s*\(\s*", "linear-gradient(");
                css = Regex.Replace(css, @"\s*\)\s*", ")");

                return $"<style>{css}</style>";
            },
            RegexOptions.Singleline | RegexOptions.IgnoreCase);
    }
}
