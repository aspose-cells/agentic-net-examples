// Title: Confirm worksheet names are rendered as <h1> headings in HTML saved with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook using Aspose.Cells, exports it to HTML, then parses the HTML to ensure each worksheet name is wrapped in an <h1> element and logs any missing headings. | Create a C# function that reads the HTML file produced by Aspose.Cells and returns true only if every sheet title appears inside an <h1> tag, otherwise returns false and lists absent headings.
// Common Searches: aspocells c# ensure sheet titles appear as h1 tags in exported html | how to validate html output contains worksheet headings after saving workbook as html with aspose.cells | c# read generated html and check for missing h1 elements for each Excel sheet | verify that Aspose.Cells HTML export includes <h1> headings for all worksheets
// Tags: Aspose.Cells HTML export worksheet heading verification | C# parse generated HTML for h1 sheet titles | validate sheet name headings in Aspose.Cells HTML output | read Aspose.Cells HTML file to check h1 tags | ensure Excel sheet titles are included as h1 in exported HTML

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, saves it as HTML, reads the generated file, and verifies that each worksheet name is present inside an <h1> tag, reporting any missing headings.
class HtmlHeadingVerifier
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path where the HTML will be saved
        string htmlPath = "output.html";

        // Load the workbook (using the provided load rule)
        Workbook workbook = new Workbook(excelPath);

        // Save the workbook as HTML (using the provided save rule)
        workbook.Save(htmlPath, SaveFormat.Html);

        // Read the generated HTML content
        string htmlContent = File.ReadAllText(htmlPath);

        // Verify that each worksheet title appears as a heading tag in the HTML
        bool allHeadingsPresent = true;

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Aspose.Cells typically renders the sheet name inside an <h1> tag
            string expectedHeading = $"<h1>{System.Web.HttpUtility.HtmlEncode(sheet.Name)}</h1>";

            if (!htmlContent.Contains(expectedHeading))
            {
                Console.WriteLine($"Missing heading for worksheet: {sheet.Name}");
                allHeadingsPresent = false;
            }
        }

        if (allHeadingsPresent)
        {
            Console.WriteLine("All worksheet titles are correctly represented as heading tags in the HTML.");
        }
        else
        {
            Console.WriteLine("One or more worksheet titles are missing heading tags in the HTML.");
        }
    }
}
