// Title: Verify HTML <img> tags are ignored and cell formatting stays unchanged when loading HTML into Aspose.Cells (C#)
// Description: C# example that loads an HTML string containing <img> tags into an Aspose.Cells workbook, confirms no EmbeddedImage objects are created, and validates that default cell styles (regular font, size 10) are preserved. The program outputs the verification results and saves the workbook.
// Keywords: Aspose.Cells HTML import ignore images | C# load HTML to Excel without images | verify cell style after HTML load | Aspose.Cells EmbeddedImage detection | HTML <img> tag handling in Aspose.Cells
// Common Searches: how to prevent images from being imported with Aspose.Cells HTML load | check if <img> tags affect cell formatting in Aspose.Cells | Aspose.Cells ignore img tag when converting HTML to Excel | verify no embedded images after loading HTML in Aspose.Cells | default cell style after HTML import Aspose.Cells
// Developer Intent: Confirm that image tags in the source HTML are not imported as embedded images and that cell formatting remains at the default settings.
// Use Cases: Load an HTML snippet containing <img> elements into a Workbook and programmatically ensure no cells contain EmbeddedImage objects. | Iterate through all cells after import to verify that the font remains regular, non‑italic, and size 10. | Save the workbook after verification to demonstrate a successful load without image artifacts.
// AI Prompts: Write C# code using Aspose.Cells to load HTML with <img> tags and assert that no cells have EmbeddedImage objects. | Create a method that checks each cell's style after loading HTML to ensure the font is not bold, not italic, and size equals 10. | Explain how HtmlLoadOptions can be configured (or left default) to ignore images when importing HTML into an Aspose.Cells workbook.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlImageIgnoreDemo
{
    // C# example that loads an HTML string containing <img> tags into an Aspose.Cells workbook, confirms no EmbeddedImage objects are created, and validates that default cell styles (regular font, size 10) are preserved. The program outputs the verification results and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Sample HTML containing an <img> tag and some text.
            string htmlContent = @"
                <html>
                    <body>
                        <p>Before image <img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUA'/> after image.</p>
                        <p>Another paragraph without images.</p>
                    </body>
                </html>";

            // Convert the HTML string to a memory stream.
            byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Create HTML load options (default options are sufficient for this test).
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                // Load the HTML content into a workbook using the provided load rule.
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Verify that no cell contains an embedded image (i.e., <img> tags are ignored).
                bool imageTagIgnored = true;
                foreach (Cell cell in cells)
                {
                    if (cell.EmbeddedImage != null)
                    {
                        imageTagIgnored = false;
                        Console.WriteLine($"Image found in cell {cell.Name} – image tags were not ignored.");
                        break;
                    }
                }

                // Additionally, confirm that cell formatting is not altered by the image tag.
                // For this simple test we expect default style (no bold, no italic, default font size).
                bool formattingUnchanged = true;
                foreach (Cell cell in cells)
                {
                    Style style = cell.GetStyle();
                    if (style.Font.IsBold || style.Font.IsItalic || style.Font.Size != 10) // default size is 10
                    {
                        formattingUnchanged = false;
                        Console.WriteLine($"Formatting changed in cell {cell.Name}.");
                        break;
                    }
                }

                // Output verification results.
                Console.WriteLine($"Image tags ignored: {imageTagIgnored}");
                Console.WriteLine($"Cell formatting unchanged: {formattingUnchanged}");

                // Save the workbook to verify that the load succeeded (uses the provided save rule).
                workbook.Save("HtmlImageIgnoreResult.xlsx");
            }
        }
    }
}
