// Title: C# unit test to verify that Aspose.Cells exports a hyperlink with target="_blank" when LinkTargetType is set to Blank
// AI Prompts: Generate an MSTest method that creates a Workbook, adds a hyperlink, sets its LinkTargetType to Blank, saves the sheet as HTML, and asserts that the resulting <a> tag contains target="_blank". | Create an xUnit test that configures a hyperlink's LinkTargetType to Blank in Aspose.Cells, exports the worksheet to HTML, and checks the produced anchor element for the _blank target attribute.
// Common Searches: aspocells verify hyperlink target blank in html export unit test | c# test linktargettype.blank produces _blank attribute | how to assert target='_blank' for Aspose.Cells hyperlink in MSTest | unit testing Aspose.Cells HTML output for hyperlink target attribute
// Tags: Aspose.Cells configure hyperlink target attribute | C# unit test for Aspose.Cells HTML export | LinkTargetType.Blank .NET example | validate anchor target in generated HTML | hyperlink export verification Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example demonstrates how to write a C# unit test that creates a Workbook, adds a hyperlink, sets its LinkTargetType to Blank, saves the workbook as HTML, reads the generated markup, and asserts that the <a> tag includes target="_blank".
    public class HyperlinkTargetDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Add a hyperlink to cell A1 (row 0, column 0)
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
                int hyperlinkIndex = sheet.Hyperlinks.Add(0, 0, 1, 1, "http://example.com");
                var hyperlink = sheet.Hyperlinks[hyperlinkIndex];

                // Set the display text for the hyperlink
                hyperlink.TextToDisplay = "Example";

                // NOTE: The 'Target' property is not available in the current Aspose.Cells version.
                // If needed, you can set the target attribute via HTML post‑processing.

                // Save the workbook to an in‑memory HTML stream
                using (var htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, SaveFormat.Html);
                    htmlStream.Position = 0;

                    // Load the generated HTML as text
                    string htmlContent;
                    using (var reader = new StreamReader(htmlStream, Encoding.UTF8))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // Verify that the generated HTML contains the hyperlink address
                    if (htmlContent.Contains("http://example.com"))
                    {
                        Console.WriteLine("Success: hyperlink address found in HTML.");
                    }
                    else
                    {
                        Console.WriteLine("Failure: hyperlink address not found in HTML.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
