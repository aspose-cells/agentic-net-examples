// Title: Generate HTML from an Aspose.Cells workbook and strip inline style attributes using C#
// AI Prompts: Create C# code that saves a Workbook to an HTML file with Aspose.Cells, then removes every style="..." attribute from the resulting markup using a regular expression. | Write a .NET method that loads an existing HTML file, parses it, and deletes all inline style attributes while preserving the rest of the document.
// Common Searches: how to remove inline CSS from Aspose.Cells generated HTML in C# | C# regex to delete style attributes after saving Excel as HTML | Aspose.Cells export to HTML without embedded style tags | post‑process HTML output from Aspose.Cells to clean up style attributes | remove style="" from HTML file using .NET
// Tags: Aspose.Cells HTML output sanitization | C# regex for HTML style removal | post‑processing generated HTML in .NET | Excel to HTML conversion without embedded CSS | clean HTML markup after workbook save

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program creates a workbook, populates cells, saves it as an HTML file via Aspose.Cells, then reads the generated HTML and uses a regular expression to strip all inline style attributes before writing the cleaned markup back to disk.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a workbook and populate it with data.
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // -------------------------------------------------
            // 2. Save the workbook as HTML.
            //    (Aspose.Cells does not provide a direct switch to
            //     disable inline styles, so we will clean them
            //     after the file is generated.)
            // -------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportHtmlAsSingleFile is not available in this version;
            // the default behavior (single file) is sufficient for this example.

            string htmlPath = "output.html";

            // Save the workbook
            workbook.Save(htmlPath, htmlOptions);

            // -------------------------------------------------
            // 3. Remove inline style attributes from the generated HTML.
            // -------------------------------------------------
            if (File.Exists(htmlPath))
            {
                try
                {
                    string htmlContent = File.ReadAllText(htmlPath);

                    // Regex to match style="...". Handles single or double quotes.
                    string cleanedHtml = Regex.Replace(
                        htmlContent,
                        @"\sstyle\s*=\s*(['""])[^'""]*\1",
                        string.Empty,
                        RegexOptions.IgnoreCase);

                    // Save the cleaned HTML back to disk
                    File.WriteAllText(htmlPath, cleanedHtml);
                    Console.WriteLine("Inline style attributes were removed.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing HTML file: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"HTML file not found at path: {htmlPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
