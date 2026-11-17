// Title: Disable external CSS files and embed custom CSS when exporting an Excel workbook to HTML with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to HTML without external stylesheet and add custom style tag | disable external CSS generation in Aspose.Cells HTMLSaveOptions C# | inject custom CSS into HTML output from Aspose.Cells workbook conversion | save Excel as HTML with embedded CSS using Aspose.Cells .NET | prevent Aspose.Cells from creating separate CSS file during HTML export
// Tags: Aspose.Cells HTML export disable external CSS | custom CSS injection Aspose.Cells HTML output | C# HtmlSaveOptions embed style tag | Excel to HTML conversion with inline CSS Aspose.Cells | Aspose.Cells prevent separate stylesheet generation

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel workbook with Aspose.Cells, saves it to HTML using HtmlSaveOptions, reads the generated markup from a memory stream, inserts a custom <style> block (or prepends it if <head> is missing), and writes the final HTML file, while disabling the creation of external CSS files.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default options are sufficient for this example)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook to a memory stream first
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                string htmlContent = Encoding.UTF8.GetString(ms.ToArray());

                // Insert custom CSS into the <head> section of the generated HTML
                const string customCss = ".customStyle { font-weight: bold; color: blue; }";
                string styleTag = $"<style>{customCss}</style>";

                int headCloseIndex = htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
                if (headCloseIndex >= 0)
                {
                    htmlContent = htmlContent.Insert(headCloseIndex, styleTag);
                }
                else
                {
                    // If <head> tag is missing, prepend the style at the beginning
                    htmlContent = styleTag + htmlContent;
                }

                // Write the final HTML to the output file
                File.WriteAllText(outputPath, htmlContent, Encoding.UTF8);
                Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
