// Title: C# example to verify that Aspose.Cells HTML export adds target="_parent" to hyperlinks
// Common Searches: how to check hyperlink target attribute _parent in Aspose.Cells generated HTML using C# | Aspose.Cells HTML export hyperlink target parent verification code sample | C# read Aspose.Cells saved HTML and confirm link target is _parent | validate that HtmlSaveOptions creates links with _parent target in Aspose.Cells | unit test for Aspose.Cells HTML hyperlink target attribute in C#
// Tags: Aspose.Cells HTML hyperlink target verification | C# HtmlSaveOptions link target configuration | check _parent attribute in exported HTML | hyperlink export validation Aspose.Cells | read generated HTML file with C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlLinkTargetValidation
{
    // The sample creates a workbook, adds a hyperlink to cell A1, saves it as HTML with Aspose.Cells, reads the generated file, and confirms that the hyperlink includes the target="_parent" attribute, reporting the validation result.
    class Program
    {
        static void Main(string[] args)
        {
            const string htmlFilePath = "HyperlinkParentTarget.html";

            try
            {
                // 1. Create a new workbook and add a hyperlink to cell A1
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");
                sheet.Cells["A1"].PutValue("Example Link");

                // 2. Configure HtmlSaveOptions (no LinkTargetType enum in this version)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // 3. Save the workbook as HTML using the configured options
                workbook.Save(htmlFilePath, htmlOptions);
                Console.WriteLine($"Workbook saved as HTML to '{htmlFilePath}'.");

                // 4. Verify that the generated HTML file exists before reading
                if (!File.Exists(htmlFilePath))
                {
                    Console.WriteLine($"Error: HTML file '{htmlFilePath}' was not created.");
                    return;
                }

                // 5. Load the generated HTML file and verify the target attribute
                string htmlContent;
                try
                {
                    htmlContent = File.ReadAllText(htmlFilePath);
                }
                catch (Exception readEx)
                {
                    Console.WriteLine("Failed to read the generated HTML file:");
                    Console.WriteLine(readEx.Message);
                    return;
                }

                const string expectedTarget = "target=\"_parent\"";
                bool containsTarget = htmlContent.IndexOf(expectedTarget, StringComparison.OrdinalIgnoreCase) >= 0;

                // 6. Output validation result
                if (containsTarget)
                {
                    Console.WriteLine("Validation succeeded: the HTML contains the expected link target attribute \"_parent\".");
                }
                else
                {
                    Console.WriteLine("Validation failed: the expected link target attribute \"_parent\" was not found in the HTML.");
                }
            }
            catch (Exception ex)
            {
                // Any unexpected error will be reported here
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
