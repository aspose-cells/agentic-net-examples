// Title: Validate Aspose.Cells HTML export adds target="_blank" when LinkTargetType is set to Blank (C#)
// Description: This C# example creates a workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank, saves the workbook as HTML, reads the output, and verifies that the generated <a> tag contains the target="_blank" attribute.
// Keywords: Aspose.Cells | HtmlSaveOptions | LinkTargetType | HtmlLinkTargetType.Blank | C# hyperlink target | validate _blank attribute | export workbook to HTML | hyperlink rendering Aspose
// Common Searches: Aspose.Cells set hyperlink target _blank | C# check target attribute in exported HTML | HtmlSaveOptions LinkTargetType Blank example | verify HTML link opens in new tab Aspose | how to validate Aspose.Cells HTML output
// Developer Intent: Confirm that the HTML file produced by Aspose.Cells contains hyperlinks with target="_blank" after configuring LinkTargetType to Blank.
// Use Cases: Automated testing to ensure exported HTML links open in a new browser tab. | Generating HTML reports where external links must open in a separate window for better UX. | Debugging hyperlink rendering issues in Aspose.Cells HTML exports.
// AI Prompts: Create a C# unit test that loads the saved HTML and asserts every <a> tag includes target="_blank" when LinkTargetType is Blank. | Provide a code snippet that parses the generated HTML, extracts all anchor elements, and prints their href and target values. | Explain the effect of each HtmlLinkTargetType option on hyperlink markup in Aspose.Cells HTML export.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlLinkTargetValidation
{
    // This C# example creates a workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank, saves the workbook as HTML, reads the output, and verifies that the generated <a> tag contains the target="_blank" attribute.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put display text into a cell and add a hyperlink to it
            worksheet.Cells["A1"].PutValue("Visit Aspose");
            worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

            // Configure HTML save options to set the link target type to "_blank"
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

            // Define the output HTML file path
            string htmlPath = "LinkTargetBlankExample.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, saveOptions);

            // Read the generated HTML file as text
            string htmlContent = File.ReadAllText(htmlPath);

            // Check if the hyperlink contains target="_blank"
            bool containsBlankTarget = htmlContent.Contains("target=\"_blank\"");

            // Output the validation result
            Console.WriteLine(containsBlankTarget
                ? "Validation succeeded: link target attribute is set to \"_blank\"."
                : "Validation failed: link target attribute \"_blank\" not found.");

            // Optional: display a snippet of the hyperlink line for debugging
            if (!containsBlankTarget)
            {
                // Find the line containing the hyperlink (simple heuristic)
                using (StringReader reader = new StringReader(htmlContent))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.Contains("<a") && line.Contains("href"))
                        {
                            Console.WriteLine("Hyperlink line: " + line.Trim());
                            break;
                        }
                    }
                }
            }
        }
    }
}
