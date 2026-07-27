// Title: C# – Export Workbook to HTML Excluding Hidden Sheets & Verify TableCssId Styles with Aspose.Cells
// Description: Demonstrates how to save a workbook that contains both visible and hidden worksheets to HTML using Aspose.Cells. The example configures HtmlSaveOptions to skip hidden sheets, generate a separate CSS file per worksheet, and remove unused styles, then reads the CSS to confirm that only TableCssId‑prefixed styles from visible sheets are present.
// Keywords: Aspose.Cells HTML export | C# export hidden worksheets | TableCssId style generation | ExportWorksheetCSSSeparately | ExcludeUnusedStyles | verify CSS content | Aspose.Cells HtmlSaveOptions | hidden sheet CSS exclusion
// Common Searches: Aspose.Cells export hidden worksheets to HTML C# | TableCssId CSS only for visible sheets Aspose | HtmlSaveOptions exclude hidden sheets example | how to check generated CSS after HTML export Aspose.Cells | C# unit test Aspose.Cells HTML export hidden sheet
// Developer Intent: Save a workbook as HTML while omitting hidden worksheets and ensure that TableCssId‑prefixed CSS is generated solely for visible sheets.
// Use Cases: Create an HTML report from a workbook that contains confidential or auxiliary data on hidden sheets, preventing that data from being exposed in the output. | Reduce CSS payload by generating per‑worksheet styles and discarding unused definitions, improving page load performance. | Programmatically validate the exported CSS to guarantee compliance with styling policies or to support automated testing pipelines.
// AI Prompts: Generate C# code using Aspose.Cells that exports a workbook to HTML, skips hidden worksheets, creates a separate CSS file, and verifies that TableCssId styles appear only for visible sheets. | Explain the impact of ExportHiddenWorksheet, ExportWorksheetCSSSeparately, and ExcludeUnusedStyles on the HTML and CSS files produced by Aspose.Cells. | Write a C# unit test that builds a workbook with a visible and a hidden sheet, saves it with the appropriate HtmlSaveOptions, and asserts that the resulting CSS contains TableCssId but not the hidden sheet's TableStyleMedium9.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Tables;

// Demonstrates how to save a workbook that contains both visible and hidden worksheets to HTML using Aspose.Cells. The example configures HtmlSaveOptions to skip hidden sheets, generate a separate CSS file per worksheet, and remove unused styles, then reads the CSS to confirm that only TableCssId‑prefixed styles from visible sheets are present.
class ExportWorkbookHiddenSheetsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Visible worksheet ----------
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";

            // Add sample data
            visibleSheet.Cells["A1"].PutValue("Header1");
            visibleSheet.Cells["B1"].PutValue("Header2");
            visibleSheet.Cells["A2"].PutValue("V1");
            visibleSheet.Cells["B2"].PutValue("V2");

            // Create a table (ListObject) and assign a style
            int visibleTableIdx = visibleSheet.ListObjects.Add(0, 0, 2, 2, true);
            ListObject visibleTable = visibleSheet.ListObjects[visibleTableIdx];
            visibleTable.TableStyleName = "TableStyleMedium2";

            // ---------- Hidden worksheet ----------
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.IsVisible = false; // Hide the sheet

            // Add sample data to hidden sheet
            hiddenSheet.Cells["A1"].PutValue("Header1");
            hiddenSheet.Cells["B1"].PutValue("Header2");
            hiddenSheet.Cells["A2"].PutValue("H1");
            hiddenSheet.Cells["B2"].PutValue("H2");

            // Create a table on the hidden sheet with a different style
            int hiddenTableIdx = hiddenSheet.ListObjects.Add(0, 0, 2, 2, true);
            ListObject hiddenTable = hiddenSheet.ListObjects[hiddenTableIdx];
            hiddenTable.TableStyleName = "TableStyleMedium9";

            // ---------- HTML save options ----------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false,          // Do NOT export hidden worksheets
                ExportWorksheetCSSSeparately = true,    // Generate CSS per worksheet
                ExcludeUnusedStyles = true              // Remove styles not used in visible sheets
            };

            // Save the workbook as HTML (creates .html and .css files)
            string htmlFilePath = "WorkbookExport.html";
            workbook.Save(htmlFilePath, saveOptions);

            // ---------- Verify generated CSS ----------
            // When ExportWorksheetCSSSeparately is true, a .css file is created alongside the .html file.
            string cssFilePath = Path.ChangeExtension(htmlFilePath, ".css");
            if (File.Exists(cssFilePath))
            {
                string cssContent = File.ReadAllText(cssFilePath);

                // TableCssId prefix is used for styles generated for visible sheets.
                bool containsTableCssId = cssContent.Contains("TableCssId");
                // The hidden sheet's style should not appear because ExportHiddenWorksheet is false.
                bool containsHiddenSheetStyle = cssContent.Contains("TableStyleMedium9");

                Console.WriteLine($"CSS contains TableCssId prefixed styles: {containsTableCssId}");
                Console.WriteLine($"CSS contains hidden sheet style (should be false): {containsHiddenSheetStyle}");
            }
            else
            {
                Console.WriteLine("CSS file was not generated.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
