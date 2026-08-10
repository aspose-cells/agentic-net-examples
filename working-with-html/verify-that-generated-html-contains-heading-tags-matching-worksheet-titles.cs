// Title: Validate worksheet titles appear as HTML heading tags when exporting to HTML with Aspose.Cells for .NET
// Description: C# sample that creates a workbook with "Sales" and "Inventory" sheets, exports all sheets to HTML using HtmlSaveOptions (ShowAllSheets), then uses regular expressions to confirm each sheet name is wrapped in an <h1>-<h6> element and prints the verification result.
// Keywords: Aspose.Cells | C# HTML export | worksheet title heading | HtmlSaveOptions ShowAllSheets | regex HTML validation | Excel to HTML conversion | multi‑sheet export | heading verification | automated testing | CI validation
// Common Searches: Aspose.Cells export workbook to HTML with sheet headings | C# check if worksheet names are in <h1> tags after HTML export | Regex verify sheet titles in Aspose.Cells generated HTML | ShowAllSheets option heading tags Aspose.Cells example | Automated test for HTML output of multi‑sheet workbook
// Developer Intent: Confirm that every worksheet name is rendered as an HTML heading element in the exported file.
// Use Cases: Automated QA for HTML reports generated from Excel workbooks, ensuring each sheet starts with a proper heading. | Creating documentation where each worksheet becomes a distinct HTML section with its name as a heading. | Integrating a validation step in CI/CD pipelines that fails the build if any worksheet title is missing from the HTML output.
// AI Prompts: Generate C# code using Aspose.Cells to export a multi‑sheet workbook to HTML and verify each sheet name appears inside an <h2> tag with Regex. | Write an NUnit test that asserts worksheet titles are present as heading tags in the HTML produced by Aspose.Cells. | Explain how to configure HtmlSaveOptions to set a custom heading level for worksheet titles during HTML export.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlHeadingVerification
{
    // C# sample that creates a workbook with "Sales" and "Inventory" sheets, exports all sheets to HTML using HtmlSaveOptions (ShowAllSheets), then uses regular expressions to confirm each sheet name is wrapped in an <h1>-<h6> element and prints the verification result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Rename the default sheet and add another sheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sales";
            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");

            // Populate some sample data in both sheets
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Amount");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(150);

            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Quantity");
            sheet2.Cells["A2"].PutValue("Screws");
            sheet2.Cells["B2"].PutValue(500);

            // Configure HTML save options (export all sheets)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Ensure each worksheet is rendered with its title
                ExportActiveWorksheetOnly = false,
                ShowAllSheets = true
            };

            // Define output HTML file path
            string htmlPath = Path.Combine(Path.GetTempPath(), "WorkbookExport.html");

            // Save the workbook as HTML
            workbook.Save(htmlPath, htmlOptions);

            // Load the generated HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Verify that each worksheet title appears inside a heading tag (e.g., <h1> or <h2>)
            bool salesHeadingFound = Regex.IsMatch(htmlContent, @"<h[1-6][^>]*>\s*Sales\s*</h[1-6]>", RegexOptions.IgnoreCase);
            bool inventoryHeadingFound = Regex.IsMatch(htmlContent, @"<h[1-6][^>]*>\s*Inventory\s*</h[1-6]>", RegexOptions.IgnoreCase);

            // Output verification results
            Console.WriteLine($"HTML file saved to: {htmlPath}");
            Console.WriteLine($"Sales sheet heading present: {salesHeadingFound}");
            Console.WriteLine($"Inventory sheet heading present: {inventoryHeadingFound}");

            // Simple assertion (optional)
            if (salesHeadingFound && inventoryHeadingFound)
            {
                Console.WriteLine("Verification succeeded: All worksheet titles are present as headings in the HTML.");
            }
            else
            {
                Console.WriteLine("Verification failed: One or more worksheet titles are missing from the HTML headings.");
            }
        }
    }
}
