// Title: Export a workbook to HTML with only visible worksheets and verify TableCssId styles are generated for visible tables using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook to HTML excluding hidden sheets | How to ensure TableCssId CSS classes are created only for visible tables in Aspose.Cells HTML output | Verify TableCssId style count after saving a workbook that contains hidden worksheets in .NET | Aspose.Cells HtmlSaveOptions SheetSet.Visible example with ListObject tables | Check HTML export of hidden worksheets using Aspose.Cells for .NET
// Tags: Aspose.Cells HTML export visible sheets | TableCssId style generation Aspose.Cells | C# hide worksheet during HTML save | ListObject table HTML export Aspose.Cells | validate table CSS ids in exported HTML

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Rendering;

namespace ExportHiddenSheetsDemo
{
    // The sample creates a workbook with a visible sheet and a hidden sheet, each containing a ListObject table. It saves the workbook to HTML using HtmlSaveOptions with SheetSet.Visible so only visible worksheets are exported. After saving, the code reads the HTML file, counts the .TableCssId style definitions, compares this count to the number of tables on visible sheets, and reports success when they match, confirming that TableCssId styles are generated solely for tables on visible worksheets.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path for the exported HTML file
                const string outputHtml = "WorkbookExport.html";

                // ------------------------------------------------------------
                // 1. Create a new workbook and add data with tables
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();

                // ----- Visible sheet -----
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";

                // Populate some data
                visibleSheet.Cells["A1"].PutValue("Product");
                visibleSheet.Cells["B1"].PutValue("Quantity");
                visibleSheet.Cells["A2"].PutValue("Apple");
                visibleSheet.Cells["B2"].PutValue(10);
                visibleSheet.Cells["A3"].PutValue("Orange");
                visibleSheet.Cells["B3"].PutValue(20);

                // Create a table (ListObject) on the visible sheet
                int visibleTableIndex = visibleSheet.ListObjects.Add(0, 0, 2, 1, true); // rows 0-2, cols 0-1, has headers
                ListObject visibleTable = visibleSheet.ListObjects[visibleTableIndex];
                // Name property may not be available in older Aspose.Cells versions; omit if not supported
                visibleTable.ShowHeaderRow = true;
                visibleTable.ShowTableStyleFirstColumn = true;
                visibleTable.ShowTableStyleLastColumn = true;
                visibleTable.ShowTableStyleRowStripes = true;
                visibleTable.ShowTableStyleColumnStripes = false;

                // ----- Hidden sheet -----
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.IsVisible = false; // Hide the sheet

                // Populate data on hidden sheet
                hiddenSheet.Cells["A1"].PutValue("Region");
                hiddenSheet.Cells["B1"].PutValue("Sales");
                hiddenSheet.Cells["A2"].PutValue("North");
                hiddenSheet.Cells["B2"].PutValue(5000);
                hiddenSheet.Cells["A3"].PutValue("South");
                hiddenSheet.Cells["B3"].PutValue(3000);

                // Create a table on the hidden sheet
                int hiddenTableIndex = hiddenSheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject hiddenTable = hiddenSheet.ListObjects[hiddenTableIndex];
                // Omit Name assignment for compatibility
                hiddenTable.ShowHeaderRow = true;
                hiddenTable.ShowTableStyleFirstColumn = true;
                hiddenTable.ShowTableStyleLastColumn = true;
                hiddenTable.ShowTableStyleRowStripes = true;
                hiddenTable.ShowTableStyleColumnStripes = false;

                // ------------------------------------------------------------
                // 2. Export the workbook to HTML (visible sheets only)
                // ------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    SheetSet = SheetSet.Visible // Export only visible sheets
                };

                workbook.Save(outputHtml, htmlOptions);
                Console.WriteLine($"Workbook exported to HTML file: {outputHtml}");

                // ------------------------------------------------------------
                // 3. Verify that TableCssId prefixed styles are generated only
                //    for the visible sheet's table.
                // ------------------------------------------------------------
                if (!File.Exists(outputHtml))
                {
                    Console.WriteLine("Error: Exported HTML file not found.");
                    return;
                }

                string htmlContent;
                try
                {
                    htmlContent = File.ReadAllText(outputHtml);
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"Failed to read HTML file: {readEx.Message}");
                    return;
                }

                // Count occurrences of TableCssId style definitions (e.g., .TableCssId0, .TableCssId1, ...)
                int styleCount = Regex.Matches(htmlContent, @"\.TableCssId\d+").Count;

                // Count tables that were actually created (visible + hidden)
                int totalTables = workbook.Worksheets[0].ListObjects.Count + workbook.Worksheets[1].ListObjects.Count;

                // Count tables on visible sheets only
                int visibleTables = 0;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.IsVisible)
                    {
                        visibleTables += ws.ListObjects.Count;
                    }
                }

                Console.WriteLine($"Total tables in workbook (including hidden sheets): {totalTables}");
                Console.WriteLine($"Tables on visible sheets: {visibleTables}");
                Console.WriteLine($"TableCssId style definitions found in HTML: {styleCount}");

                if (styleCount == visibleTables)
                {
                    Console.WriteLine("Success: TableCssId styles are generated only for visible sheets.");
                }
                else
                {
                    Console.WriteLine("Failure: Mismatch between visible tables and generated TableCssId styles.");
                }

                // Optional: Output the style definitions for manual inspection
                Console.WriteLine("\n--- Extracted TableCssId style definitions ---");
                foreach (Match match in Regex.Matches(htmlContent, @"\.TableCssId\d+\s*\{[^}]*\}"))
                {
                    Console.WriteLine(match.Value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
