// Title: C# – Apply a Custom Style to a Range and Export HTML without Unused Styles using Aspose.Cells
// Description: Demonstrates how to create a workbook with Aspose.Cells, define a custom bold blue style, apply it to cells A1:C3, add an unused style, and save the sheet as a single HTML file with HtmlSaveOptions.ExcludeUnusedStyles set to true. The example also shows how to read the generated HTML to confirm that the unused style is omitted.
// Keywords: Aspose.Cells C# | custom style range | HtmlSaveOptions ExcludeUnusedStyles | export HTML without unused CSS | Aspose.Cells HTML export example | verify style not in HTML | single file HTML Aspose.Cells
// Common Searches: Aspose.Cells apply style to range C# | ExcludeUnusedStyles HTML Aspose.Cells .NET | Save workbook as single HTML file Aspose.Cells | remove unused CSS when exporting Excel to HTML | how to verify unused style is not in exported HTML
// Developer Intent: Create a workbook, apply a custom style to a specific range, and generate HTML that contains only the styles actually used.
// Use Cases: Produce clean, lightweight HTML reports from Excel data. | Minimize HTML file size by excluding unused CSS definitions. | Automated testing to ensure compliance with style‑usage policies.
// AI Prompts: Generate C# code with Aspose.Cells that defines a bold blue style, applies it to A1:C3, creates an unused style, and saves the workbook as a single HTML file with ExcludeUnusedStyles enabled. | Explain the purpose of HtmlSaveOptions.ExcludeUnusedStyles in Aspose.Cells and provide a snippet that checks the resulting HTML for the presence of an unused style.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCustomStyleHtmlDemo
{
    // Demonstrates how to create a workbook with Aspose.Cells, define a custom bold blue style, apply it to cells A1:C3, add an unused style, and save the sheet as a single HTML file with HtmlSaveOptions.ExcludeUnusedStyles set to true. The example also shows how to read the generated HTML to confirm that the unused style is omitted.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Create a custom style that will be applied
                // -------------------------------------------------
                Style usedStyle = workbook.CreateStyle(); // Workbook.CreateStyle()
                usedStyle.Font.Name = "Arial";
                usedStyle.Font.Size = 12;
                usedStyle.Font.IsBold = true;
                usedStyle.Font.Color = Color.Blue;
                usedStyle.ForegroundColor = Color.LightYellow;
                usedStyle.Pattern = BackgroundType.Solid;

                // -------------------------------------------------
                // 2. Apply the style to a specific range (A1:C3)
                // -------------------------------------------------
                AsposeRange range = sheet.Cells.CreateRange("A1:C3"); // Cells.CreateRange(string)
                range.SetStyle(usedStyle); // Range.SetStyle(Style)

                // -------------------------------------------------
                // 3. Create another style but DO NOT apply it (unused)
                // -------------------------------------------------
                Style unusedStyle = workbook.CreateStyle(); // Workbook.CreateStyle()
                unusedStyle.Font.Name = "Times New Roman";
                unusedStyle.Font.Size = 14;
                unusedStyle.Font.Color = Color.Red;
                unusedStyle.Name = "UnusedStyle"; // give it a name for identification

                // -------------------------------------------------
                // 4. Configure HTML save options to exclude unused styles
                // -------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(); // new HtmlSaveOptions()
                htmlOptions.ExcludeUnusedStyles = true; // HtmlSaveOptions.ExcludeUnusedStyles
                htmlOptions.SaveAsSingleFile = true; // optional single file output

                // -------------------------------------------------
                // 5. Save the workbook as HTML
                // -------------------------------------------------
                string htmlPath = "StyledOutput.html";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(htmlPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(htmlPath, htmlOptions); // Workbook.Save(string, SaveOptions)

                // -------------------------------------------------
                // 6. Verify that the unused style does NOT appear in the HTML
                // -------------------------------------------------
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    bool unusedStyleFound = htmlContent.Contains("UnusedStyle");

                    Console.WriteLine("HTML saved to: " + Path.GetFullPath(htmlPath));
                    Console.WriteLine("Unused style present in HTML? " + (unusedStyleFound ? "Yes" : "No"));
                    // Expected output: No
                }
                else
                {
                    Console.WriteLine("Failed to create HTML file at: " + Path.GetFullPath(htmlPath));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
