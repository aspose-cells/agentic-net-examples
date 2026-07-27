// Title: Validate Aspose.Cells HTML Export Sets hyperlink target="_blank" via LinkTargetType (C#)
// Description: A C# example that creates a workbook, adds a hyperlink, configures HtmlSaveOptions.LinkTargetType to Blank, saves the file as HTML, reads the output, and verifies that the generated <a> tag contains target="_blank". The program prints a pass/fail message and shows a snippet of the HTML.
// Keywords: Aspose.Cells | HtmlSaveOptions | LinkTargetType | Blank | C# | hyperlink target | target=_blank | HTML export validation | Excel to HTML | Aspose.Cells example
// Common Searches: Aspose.Cells set hyperlink target blank | C# verify target=_blank in exported HTML | HtmlSaveOptions LinkTargetType example | How to make links open in new tab with Aspose.Cells | Validate Aspose.Cells HTML output hyperlink
// Developer Intent: Confirm that setting HtmlSaveOptions.LinkTargetType to Blank makes all exported hyperlinks include target="_blank".
// Use Cases: Generate HTML reports from Excel where external links must open in a new browser tab. | Automated regression test to ensure link behavior remains consistent after library updates. | Create web‑ready documentation from workbooks with uniform hyperlink targeting.
// AI Prompts: Write a C# unit test using Aspose.Cells that asserts the saved HTML contains target="_blank" on all <a> elements. | Provide a C# snippet that parses an HTML file, extracts every hyperlink, and lists its target attribute values. | Explain the impact of HtmlSaveOptions.LinkTargetType on hyperlink rendering in Aspose.Cells HTML output.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlLinkTargetValidation
{
    // A C# example that creates a workbook, adds a hyperlink, configures HtmlSaveOptions.LinkTargetType to Blank, saves the file as HTML, reads the output, and verifies that the generated <a> tag contains target="_blank". The program prints a pass/fail message and shows a snippet of the HTML.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put display text into a cell and add a hyperlink to it
                worksheet.Cells["A1"].PutValue("Visit Aspose");
                worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

                // Configure HTML save options to set the link target type to "_blank"
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Set the desired target type for hyperlinks
                    LinkTargetType = HtmlLinkTargetType.Blank
                };

                // Define the output HTML file path
                string htmlPath = "LinkTargetBlankExample.html";

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, saveOptions);

                // Ensure the HTML file was created before attempting to read it
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: The file \"{htmlPath}\" was not found.");
                    return;
                }

                // Load the generated HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // Validate that the hyperlink contains target="_blank"
                bool containsBlankTarget = htmlContent.Contains("target=\"_blank\"");

                // Output the validation result
                Console.WriteLine(containsBlankTarget
                    ? "Validation passed: link target attribute is set to \"_blank\"."
                    : "Validation failed: link target attribute \"_blank\" not found.");

                // Optional: display a snippet of the HTML for verification
                Console.WriteLine("\nHTML snippet:");
                int start = htmlContent.IndexOf("<a ", StringComparison.OrdinalIgnoreCase);
                if (start >= 0)
                {
                    int end = htmlContent.IndexOf("</a>", start, StringComparison.OrdinalIgnoreCase);
                    if (end >= 0)
                    {
                        end += 4; // Include the closing tag length
                        int length = Math.Min(200, end - start);
                        Console.WriteLine(htmlContent.Substring(start, length));
                    }
                    else
                    {
                        Console.WriteLine("Closing </a> tag not found.");
                    }
                }
                else
                {
                    Console.WriteLine("<a> tag not found in the generated HTML.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
