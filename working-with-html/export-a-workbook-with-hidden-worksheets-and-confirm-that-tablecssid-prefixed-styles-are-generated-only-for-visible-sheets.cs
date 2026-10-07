// Title: Export a workbook to HTML with Aspose.Cells .NET while skipping hidden worksheets and confirming TableCssId styles are generated only for visible sheets
// AI Prompts: Write C# code that creates a workbook with one visible and one hidden worksheet, adds a ListObject to each, saves the workbook as HTML, and checks that only a single .TableCssId style block appears in the output. | Modify the HTML save options in Aspose.Cells to include hidden worksheets in the export but still suppress TableCssId CSS generation for those hidden sheets. | Extend the example by adding a second visible worksheet with its own table and verify that a distinct .TableCssId style block is produced for each visible sheet.
// Common Searches: Aspose.Cells .NET export only visible worksheets to HTML | C# verify TableCssId CSS blocks after saving workbook as HTML with Aspose.Cells | prevent hidden sheet CSS generation when exporting workbook to HTML using Aspose.Cells | how to count TableCssId style definitions in HTML output from Aspose.Cells | HTMLSaveOptions default handling of hidden worksheets Aspose.Cells
// Tags: HTML export hidden worksheets Aspose.Cells | TableCssId CSS generation visible sheets | C# ListObject to HTML Aspose.Cells | validate exported HTML style blocks Aspose.Cells | hide worksheet before HTML save Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The program creates a workbook with a visible sheet and a hidden sheet, adds a ListObject (table) to each, hides the second sheet, exports the workbook to HTML using default options (which omit hidden sheets), reads the generated HTML, counts .TableCssId style blocks with a regex, and confirms that only the visible sheet's table produced a CSS block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // ----- Visible worksheet -----
            Worksheet wsVisible = wb.Worksheets[0];
            wsVisible.Name = "VisibleSheet";

            // Populate data
            wsVisible.Cells["A1"].PutValue("Header1");
            wsVisible.Cells["B1"].PutValue("Header2");
            wsVisible.Cells["A2"].PutValue(1);
            wsVisible.Cells["B2"].PutValue(2);

            // Add a table (list object) to the visible sheet
            int visibleTableIndex = wsVisible.ListObjects.Add(0, 0, 2, 2, true);
            ListObject visibleTable = wsVisible.ListObjects[visibleTableIndex];
            // Set display name (compatible with all Aspose.Cells versions)
            visibleTable.DisplayName = "VisibleTable";

            // ----- Hidden worksheet -----
            Worksheet wsHidden = wb.Worksheets.Add("HiddenSheet");

            // Populate data
            wsHidden.Cells["A1"].PutValue("Header1");
            wsHidden.Cells["B1"].PutValue("Header2");
            wsHidden.Cells["A2"].PutValue(10);
            wsHidden.Cells["B2"].PutValue(20);

            // Add a table to the hidden sheet
            int hiddenTableIndex = wsHidden.ListObjects.Add(0, 0, 2, 2, true);
            ListObject hiddenTable = wsHidden.ListObjects[hiddenTableIndex];
            hiddenTable.DisplayName = "HiddenTable";

            // Hide the second worksheet
            wsHidden.IsVisible = false;

            // Export the workbook to HTML
            string htmlPath = "ExportedWorkbook.html";
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            wb.Save(htmlPath, htmlOptions); // Only visible sheets are exported by default

            // Ensure the HTML file was created before reading
            if (File.Exists(htmlPath))
            {
                // Load the generated HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // Find all TableCssId style blocks (e.g., .TableCssId0 { ... })
                MatchCollection styleMatches = Regex.Matches(htmlContent, @"\.TableCssId\d+\s*\{[^}]*\}");

                Console.WriteLine($"Total TableCssId style blocks found: {styleMatches.Count}");

                // Expect only one style block (from the visible sheet)
                if (styleMatches.Count == 1)
                {
                    Console.WriteLine("Success: TableCssId styles generated only for visible worksheets.");
                }
                else
                {
                    Console.WriteLine("Failure: TableCssId styles were generated for hidden worksheets.");
                }
            }
            else
            {
                Console.WriteLine($"Error: HTML file '{htmlPath}' was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An exception occurred: {ex.Message}");
        }
    }
}
