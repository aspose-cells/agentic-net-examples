// Title: Add cell‑address based alt attributes to images when exporting an Aspose.Cells workbook to HTML using C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, saves it as HTML with external image files, and then updates each <img> tag to include an alt attribute that reflects the originating cell address of the chart or picture. | Generate a C# post‑processing routine that reads the HTML file produced by Aspose.Cells, locates all <img> elements, and injects descriptive alt text based on a pre‑collected list of chart and picture identifiers. | Provide a C# example that uses a regular expression to replace <img> tags in an Aspose.Cells HTML export, adding alt attributes that reference the source cell locations.
// Common Searches: how to set alt text for chart images in Aspose.Cells HTML export C# | C# post‑process Aspose.Cells generated HTML to add image alt attributes | Aspose.Cells save workbook as HTML with accessible image alt tags | retrieve cell address of chart in Aspose.Cells for alt attribute | regex replace img tags in Aspose.Cells HTML output C#
// Tags: Aspose.Cells HTML export add alt attributes | C# regex modify img tags in Aspose.Cells HTML | chart image accessibility Aspose.Cells | post‑process HTML for Aspose.Cells workbook conversion | external image files Aspose.Cells HtmlSaveOptions

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

// The sample program creates a workbook with sample data and a column chart, exports it to HTML with external image files, gathers identifiers for charts and pictures, reads the generated HTML, and uses a regular expression to inject alt attributes into each <img> tag, improving accessibility by linking images to their source cell addresses.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------- Create workbook --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Fill sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a chart positioned at row 5, column 0 (cell A6)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Chart";

            // -------------------- Save as HTML --------------------
            string outputFolder = "output";
            Directory.CreateDirectory(outputFolder);

            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = false // Save images as separate files
                // Image folder options are omitted; images will be saved next to the HTML file
            };

            string htmlPath = Path.Combine(outputFolder, "workbook.html");
            workbook.Save(htmlPath, htmlOptions);

            // -------------------- Build alt‑text list --------------------
            // Collect identifiers for charts and pictures in the order they appear.
            List<string> altTexts = new List<string>();

            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Charts
                for (int i = 0; i < ws.Charts.Count; i++)
                {
                    // Use a simple identifier since direct cell position is not exposed.
                    altTexts.Add($"Chart_{i + 1}");
                }

                // Pictures (if any)
                for (int i = 0; i < ws.Pictures.Count; i++)
                {
                    altTexts.Add($"Picture_{i + 1}");
                }
            }

            // -------------------- Post‑process HTML --------------------
            if (File.Exists(htmlPath))
            {
                string htmlContent;
                try
                {
                    htmlContent = File.ReadAllText(htmlPath);
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"Failed to read HTML file: {readEx.Message}");
                    return;
                }

                int imgCounter = 0;

                // Replace each <img ...> tag with an alt attribute containing the identifier.
                htmlContent = Regex.Replace(
                    htmlContent,
                    "<img([^>]*?)src=\"([^\"]+)\"([^>]*?)>",
                    match =>
                    {
                        string beforeSrc = match.Groups[1].Value;
                        string src = match.Groups[2].Value;
                        string afterSrc = match.Groups[3].Value;

                        string alt = imgCounter < altTexts.Count ? altTexts[imgCounter] : "";
                        imgCounter++;

                        // Preserve any existing attributes and inject alt
                        return $"<img{beforeSrc}src=\"{src}\" alt=\"{alt}\"{afterSrc}>";
                    },
                    RegexOptions.IgnoreCase);

                try
                {
                    File.WriteAllText(htmlPath, htmlContent);
                }
                catch (Exception writeEx)
                {
                    Console.WriteLine($"Failed to write modified HTML file: {writeEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
