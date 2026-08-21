// Title: Export Excel to Compact HTML with Comments using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a cell comment, and save it as compact HTML with minimal CSS by enabling ExcludeUnusedStyles, including comments with IsExportComments, and generating separate files via SaveAsSingleFile = false.
// Keywords: Aspose.Cells HTML export | ExcludeUnusedStyles | IsExportComments | compact HTML output | SaveAsSingleFile false | C# Excel to HTML | cell comments in HTML | Aspose.Cells HtmlSaveOptions
// Common Searches: Aspose.Cells export HTML with comments C# | How to exclude unused styles when saving Excel as HTML | Generate lightweight HTML from workbook using Aspose.Cells | Save Excel as multiple HTML files Aspose.Cells | Include cell comments in HTML export Aspose.Cells
// Developer Intent: Produce an HTML representation of a workbook that retains cell comments while stripping unused CSS and avoiding a single‑file bundle.
// Use Cases: Web preview of spreadsheets with visible comments and reduced file size. | Documentation generation where only necessary styles are embedded. | Creating separate HTML pages per worksheet for modular web deployment.
// AI Prompts: Show C# code that saves an Aspose.Cells workbook to HTML with ExcludeUnusedStyles and IsExportComments enabled. | Provide an example of exporting Excel to compact multi‑file HTML while preserving cell comments using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, add a cell comment, and save it as compact HTML with minimal CSS by enabling ExcludeUnusedStyles, including comments with IsExportComments, and generating separate files via SaveAsSingleFile = false.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment";

            // Configure HTML save options:
            // - ExcludeUnusedStyles = true (default) to keep HTML compact
            // - IsExportComments = true to include comments in the output
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExcludeUnusedStyles = true,
                IsExportComments = true,
                // Ensure the output is not saved as a single file (default behavior)
                SaveAsSingleFile = false
            };

            // Save the workbook as HTML
            string outputPath = "output.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to {outputPath} with compact styles and comments exported.");
        }
    }
}
