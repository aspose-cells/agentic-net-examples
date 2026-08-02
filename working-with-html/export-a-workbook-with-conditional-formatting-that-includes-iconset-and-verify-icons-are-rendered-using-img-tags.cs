// Title: Export IconSet Conditional Formatting to HTML with <img> Tags using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply a TrafficLights31 IconSet to cells A1:A5, configure HtmlSaveOptions to write icons as external image files, save the workbook as HTML, and verify the generated <img> tags with a regular expression.
// Keywords: Aspose.Cells | C# | IconSet export | HTML save options | conditional formatting icons | TrafficLights31 | ExportImagesAsBase64 false | attached files directory | image verification regex | web report icons
// Common Searches: Aspose.Cells export IconSet to HTML | HTML export conditional formatting icons Aspose | save Excel icons as separate image files .NET | verify <img> tags in exported HTML Aspose.Cells | configure HtmlSaveOptions for external images
// Developer Intent: Generate an HTML file from a workbook that contains an IconSet conditional format and ensure the icons are saved as external image files referenced by <img> tags.
// Use Cases: Render traffic‑light icons in a web‑based dashboard without base64 encoding. | Automate validation of exported HTML to confirm the correct number of icon images. | Integrate Excel‑derived reports into a CMS that requires images served from a specific URL path.
// AI Prompts: Write C# code with Aspose.Cells that adds a TrafficLights31 IconSet to a range and exports the workbook to HTML, saving icons as separate PNG files in a custom folder. | Provide a method that parses the saved HTML file and returns all <img> src attributes that correspond to conditional‑formatting icons. | Explain how to set HtmlSaveOptions so icons are not embedded as base64 but written to an attached files directory with a URL prefix.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsIconSetHtmlExport
{
    // Demonstrates how to create a workbook, apply a TrafficLights31 IconSet to cells A1:A5, configure HtmlSaveOptions to write icons as external image files, save the workbook as HTML, and verify the generated <img> tags with a regular expression.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample data in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue((i + 1) * 10); // 10, 20, 30, 40, 50
            }

            // 3. Add an IconSet conditional formatting to the range A1:A5
            int cfIndex = sheet.ConditionalFormattings.Add();                     // add a new conditional formatting collection
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex]; // get the collection

            // Define the cell area for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add an IconSet condition
            int conditionIndex = fcc.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcc[conditionIndex];

            // Configure the IconSet (use TrafficLights31 as an example)
            condition.IconSet.Type = IconSetType.TrafficLights31;
            condition.IconSet.ShowValue = true; // display the cell value alongside the icon

            // 4. Prepare HTML save options to ensure icons are exported as separate image files
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.ExportImagesAsBase64 = false; // export images as files, not base64 strings
            string attachedDir = "IconImages";
            Directory.CreateDirectory(attachedDir);
            htmlOptions.AttachedFilesDirectory = attachedDir; // folder where images will be saved
            htmlOptions.AttachedFilesUrlPrefix = attachedDir + "/"; // URL prefix used in the HTML <img> tags

            // 5. Save the workbook as HTML
            string htmlPath = "IconSetExport.html";
            workbook.Save(htmlPath, htmlOptions);

            // 6. Verify that the generated HTML contains <img> tags for the icons
            string htmlContent = File.ReadAllText(htmlPath);
            MatchCollection imgMatches = Regex.Matches(htmlContent, @"<img\s+[^>]*src\s*=\s*[""'][^""']+[""'][^>]*>", RegexOptions.IgnoreCase);

            Console.WriteLine($"HTML file saved to: {Path.GetFullPath(htmlPath)}");
            Console.WriteLine($"Number of <img> tags found: {imgMatches.Count}");

            // List the src attributes of the found <img> tags (first few if many)
            int displayCount = Math.Min(5, imgMatches.Count);
            for (int i = 0; i < displayCount; i++)
            {
                string imgTag = imgMatches[i].Value;
                Console.WriteLine($"Img tag {i + 1}: {imgTag}");
            }

            // 7. Optionally, list the exported image files to confirm they exist
            if (Directory.Exists(attachedDir))
            {
                string[] imageFiles = Directory.GetFiles(attachedDir);
                Console.WriteLine($"Exported image files in '{attachedDir}':");
                foreach (string file in imageFiles)
                {
                    Console.WriteLine($" - {Path.GetFileName(file)} ({new FileInfo(file).Length} bytes)");
                }
            }
        }
    }
}
