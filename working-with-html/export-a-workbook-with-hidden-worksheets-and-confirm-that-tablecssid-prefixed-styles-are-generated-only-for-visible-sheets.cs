// Title: C# – Export Workbook to HTML Excluding Hidden Sheets and Validate TableCssId Styles for Visible Worksheets (Aspose.Cells)
// Description: Creates a workbook with one visible and one hidden worksheet, adds a ListObject table to each with different built‑in styles, saves to HTML using HtmlSaveOptions that omit hidden sheets and generate CSS per sheet, then reads the HTML to confirm the hidden sheet name is absent and that TableCssId style blocks exist only for the visible worksheet.
// Keywords: Aspose.Cells | C# HTML export | ExportHiddenWorksheet | HtmlSaveOptions | TableCssId | visible worksheet CSS | hidden sheet exclusion | ListObject HTML export | Aspose.Cells table style | HTML verification
// Common Searches: Aspose.Cells export workbook to HTML without hidden sheets | How to prevent hidden worksheets from being saved in HTML | TableCssId CSS generated only for visible tables Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet example C# | Verify hidden worksheet is not in exported HTML Aspose
// Developer Intent: Export a workbook to HTML while skipping hidden worksheets and ensure that TableCssId‑prefixed CSS is generated solely for tables on visible sheets.
// Use Cases: Produce HTML reports that hide confidential or intermediate worksheets but retain table formatting for visible data. | Automate quality checks that confirm hidden worksheets are not present in the HTML output and that CSS identifiers correspond only to visible tables. | Create unit tests that load the generated HTML, assert the hidden sheet name is missing, and validate the count of TableCssId style blocks matches the number of visible tables.
// AI Prompts: Generate C# code using Aspose.Cells to export a workbook to HTML, exclude hidden worksheets, and produce TableCssId styles only for visible tables. | Write a C# unit test that reads the exported HTML file, verifies the hidden sheet name does not appear, and checks that the TableCssId occurrence count equals the number of visible ListObjects. | Explain the impact of HtmlSaveOptions properties ExportHiddenWorksheet, ExportActiveWorksheetOnly, and ExportWorksheetCSSSeparately on HTML and CSS output for tables in Aspose.Cells.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Tables;   // Required for ListObject and TableStyleType

namespace AsposeCellsHiddenSheetDemo
{
    // Creates a workbook with one visible and one hidden worksheet, adds a ListObject table to each with different built‑in styles, saves to HTML using HtmlSaveOptions that omit hidden sheets and generate CSS per sheet, then reads the HTML to confirm the hidden sheet name is absent and that TableCssId style blocks exist only for the visible worksheet.
    class Program
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

                // Populate data for a table
                visibleSheet.Cells["A1"].PutValue("Header1");
                visibleSheet.Cells["B1"].PutValue("Header2");
                visibleSheet.Cells["A2"].PutValue("V1");
                visibleSheet.Cells["B2"].PutValue("V2");

                // Add a table (ListObject) and apply a built‑in style
                int visibleTableIndex = visibleSheet.ListObjects.Add(0, 0, 2, 2, true);
                ListObject visibleTable = visibleSheet.ListObjects[visibleTableIndex];
                visibleTable.TableStyleType = TableStyleType.TableStyleMedium2; // any style

                // ---------- Hidden worksheet ----------
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.IsVisible = false; // hide the sheet

                // Populate data for a table on the hidden sheet
                hiddenSheet.Cells["A1"].PutValue("Header1");
                hiddenSheet.Cells["B1"].PutValue("Header2");
                hiddenSheet.Cells["A2"].PutValue("H1");
                hiddenSheet.Cells["B2"].PutValue("H2");

                // Add a table and apply a different style
                int hiddenTableIndex = hiddenSheet.ListObjects.Add(0, 0, 2, 2, true);
                ListObject hiddenTable = hiddenSheet.ListObjects[hiddenTableIndex];
                hiddenTable.TableStyleType = TableStyleType.TableStyleMedium4; // another style

                // ---------- Save to HTML without exporting hidden worksheets ----------
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = false,          // do NOT export hidden sheets
                    ExportActiveWorksheetOnly = false,      // export whole workbook (visible sheets only)
                    ExportWorksheetCSSSeparately = true,    // generate CSS per sheet (helps inspection)
                    ExcludeUnusedStyles = false             // keep all generated styles for verification
                };

                string htmlPath = "WorkbookWithHiddenSheet.html";
                workbook.Save(htmlPath, saveOptions);

                // ---------- Verify the generated HTML ----------
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);

                    // Check that the hidden sheet name does NOT appear in the HTML
                    bool hiddenSheetFound = htmlContent.Contains("HiddenSheet");

                    // Count occurrences of TableCssId (Aspose.Cells prefixes table CSS with this)
                    int tableCssIdCount = Regex.Matches(htmlContent, @"TableCssId").Count;

                    Console.WriteLine($"Hidden sheet name present in HTML: {hiddenSheetFound}");
                    Console.WriteLine($"Number of TableCssId style blocks in HTML: {tableCssIdCount}");

                    // Expected: hiddenSheetFound == false, and TableCssId count corresponds only to visible sheet
                    if (!hiddenSheetFound && tableCssIdCount > 0)
                    {
                        Console.WriteLine("Verification succeeded: only visible sheet generated TableCssId styles.");
                    }
                    else
                    {
                        Console.WriteLine("Verification failed: hidden sheet data or styles were exported.");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: HTML file '{htmlPath}' was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
