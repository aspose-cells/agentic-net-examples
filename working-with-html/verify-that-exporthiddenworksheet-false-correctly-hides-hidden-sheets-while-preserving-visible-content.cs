// Title: Aspose.Cells .NET – ExportHiddenWorksheet=false hides hidden worksheets in HTML export
// Description: Demonstrates creating a workbook with one visible and one hidden worksheet, then saving to HTML twice: first with HtmlSaveOptions.ExportHiddenWorksheet set to false (hidden sheet omitted) and then with it set to true (hidden sheet included).
// Keywords: Aspose.Cells HTML export | ExportHiddenWorksheet false | hide hidden worksheets .NET | HtmlSaveOptions ExportHiddenWorksheet | Aspose.Cells C# example | exclude hidden sheets HTML | include hidden sheets HTML
// Common Searches: Aspose.Cells ExportHiddenWorksheet example C# | how to hide hidden worksheets when exporting to HTML | exclude hidden sheets Aspose.Cells HTMLSaveOptions | include hidden worksheets in HTML export Aspose.Cells | Aspose.Cells HTML export hidden worksheet property
// Developer Intent: Confirm that setting ExportHiddenWorksheet to false prevents hidden worksheets from appearing in the generated HTML while all visible worksheets are fully rendered.
// Use Cases: Produce an HTML report that shows only data from visible tabs, keeping hidden tabs private. | Generate two versions of an HTML export—one without hidden sheets for public distribution and one with hidden sheets for internal review. | Programmatically toggle ExportHiddenWorksheet to control visibility of hidden worksheets in automated reporting pipelines.
// AI Prompts: Create a C# unit test that loads the two HTML files and verifies that "Hidden Data" is missing in the file generated with ExportHiddenWorksheet=false and present when true. | Write a step‑by‑step guide to validate HTML output from Aspose.Cells, including parsing the file and checking for content from hidden worksheets. | Suggest code modifications to log the names of worksheets that were exported based on the ExportHiddenWorksheet setting.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHiddenWorksheetDemo
{
    // Demonstrates creating a workbook with one visible and one hidden worksheet, then saving to HTML twice: first with HtmlSaveOptions.ExportHiddenWorksheet set to false (hidden sheet omitted) and then with it set to true (hidden sheet included).
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ----- Visible worksheet -----
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");

            // ----- Hidden worksheet -----
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            // Hide the worksheet
            hiddenSheet.IsVisible = false;

            // Configure HTML save options to exclude hidden worksheets
            HtmlSaveOptions optionsExcludeHidden = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false, // Do not export hidden sheets
                ExportActiveWorksheetOnly = false // Export the whole workbook (except hidden sheets)
            };

            // Save workbook without hidden worksheet
            workbook.Save("output_without_hidden.html", optionsExcludeHidden);

            // Change option to include hidden worksheets
            optionsExcludeHidden.ExportHiddenWorksheet = true;

            // Save workbook with hidden worksheet included
            workbook.Save("output_with_hidden.html", optionsExcludeHidden);

            Console.WriteLine("Export completed. Check the generated HTML files.");
        }
    }
}
