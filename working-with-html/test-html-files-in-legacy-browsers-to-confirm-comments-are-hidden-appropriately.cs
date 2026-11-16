// Title: Validate that Aspose.Cells for .NET hides cell comments when exporting an Excel worksheet to HTML for legacy browsers
// Common Searches: how to ensure cell comments are not included in HTML output using Aspose.Cells C# | C# code to test that exported HTML from Excel does not contain <!-- comment --> tags | verify Aspose.Cells HTML export compatibility with older browsers by checking comment visibility | sample program to detect HTML comment markers after saving workbook as HTML with Aspose.Cells | prevent Aspose.Cells from writing HTML comment nodes when converting Excel to HTML
// Tags: Aspose.Cells HTML export hide cell comments | C# validate absence of HTML comment markers | legacy browser compatibility Aspose.Cells HTML5 output | HtmlSaveOptions ExportActiveWorksheetOnly usage | Aspose.Cells workbook to HTML without comments | test generated HTML for comment tags in .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, saves it as HTML using HtmlSaveOptions (HTML5 and ExportActiveWorksheetOnly), then reads the generated file to confirm that no <!-- ... --> comment markers are present, reporting pass/fail for legacy‑browser compatibility.
class HtmlCommentVisibilityTest
{
    static void Main()
    {
        // Paths for input Excel file and output HTML file
        string excelPath = "input.xlsx";
        string htmlPath = "output.html";

        try
        {
            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Use HTML5 (default) – suitable for modern browsers
                HtmlVersion = HtmlVersion.Html5,

                // Export only the active worksheet (optional)
                ExportActiveWorksheetOnly = true
                // Note: ExportCellComments property is not available in this version of Aspose.Cells.
                // By default, cell comments are not exported to HTML.
            };

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while converting Excel to HTML: {ex.Message}");
            return;
        }

        // Verify that the HTML file was created
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: Failed to create the HTML file \"{htmlPath}\".");
            return;
        }

        try
        {
            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Check for HTML comment markers (<!-- ... -->)
            bool commentsFound = htmlContent.Contains("<!--");

            // Output the test result
            if (commentsFound)
                Console.WriteLine("Test Failed: Comments are present in the HTML output.");
            else
                Console.WriteLine("Test Passed: Comments are hidden in the HTML output.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while reading the HTML file: {ex.Message}");
        }
    }
}
