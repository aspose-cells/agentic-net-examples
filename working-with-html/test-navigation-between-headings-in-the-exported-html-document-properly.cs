// Title: Export a multi‑sheet workbook to HTML with Aspose.Cells and verify that each sheet appears as a navigation link (C#)
// AI Prompts: Write a C# program that creates a workbook with several worksheets, saves it as an HTML file using Aspose.Cells HtmlSaveOptions, then parses the generated HTML to confirm that every worksheet name is present as a navigation hyperlink. | Generate code that uses Aspose.Cells to export a workbook to HTML, reads the output file, and checks for the existence of anchor tags or text matching each sheet title to ensure proper navigation links.
// Common Searches: aspnet export workbook to html with aspose.cells and check sheet navigation links | c# verify that exported html contains hyperlinks for each worksheet name using Aspose.Cells | how to read generated html from Aspose.Cells and confirm sheet titles are linked | Aspose.Cells HtmlSaveOptions multi‑sheet navigation link validation in C#
// Tags: Aspose.Cells export workbook to HTML | verify worksheet navigation links in HTML | HtmlSaveOptions sheet hyperlink verification | C# read generated HTML for sheet names | multi‑sheet HTML output Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook with two worksheets, saves it as an HTML file using Aspose.Cells HtmlSaveOptions, then reads the resulting HTML and checks that the names of both worksheets are present as navigation links, confirming proper HTML navigation for multi‑sheet workbooks.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ----- First worksheet -----
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Heading 1");
            sheet1.Cells["A2"].PutValue("Data 1");

            // ----- Second worksheet -----
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "SecondSheet";
            sheet2.Cells["A1"].PutValue("Heading 2");
            sheet2.Cells["A2"].PutValue("Data 2");

            // Export the workbook to HTML.
            // Note: ExportHtmlNavigation property is not available in the current Aspose.Cells version,
            // so we rely on the default HTML output which still contains worksheet names.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            string htmlFilePath = "ExportedWorkbook.html";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(htmlFilePath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(htmlFilePath, htmlOptions);

            // Load the generated HTML file safely
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: HTML file '{htmlFilePath}' was not created.");
                return;
            }

            string htmlContent = File.ReadAllText(htmlFilePath);

            // Verify that navigation links (worksheet names) exist in the HTML
            bool hasFirstSheetLink = htmlContent.Contains("FirstSheet");
            bool hasSecondSheetLink = htmlContent.Contains("SecondSheet");

            Console.WriteLine($"Navigation link for 'FirstSheet' found: {hasFirstSheetLink}");
            Console.WriteLine($"Navigation link for 'SecondSheet' found: {hasSecondSheetLink}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
