// Title: Batch update linked shapes across worksheets after changing external data sources – Aspose.Cells for .NET
// Description: Loads a primary workbook, replaces its external links with new source workbooks using UpdateLinkedDataSource, refreshes every linked shape (dropdowns, list boxes, etc.) on all worksheets via Shapes.UpdateSelectedValue, optionally updates dependent pivot tables and charts with RefreshAll, and saves the result.
// Keywords: Aspose.Cells | C# | UpdateLinkedDataSource | Shapes.UpdateSelectedValue | linked shapes refresh | external workbook links | batch update workbook | refresh dropdown lists | refresh list boxes | refresh charts and pivots | automate Excel workbook | Excel automation .NET
// Common Searches: Aspose.Cells refresh linked shapes after changing source workbook | C# batch update external links and linked dropdowns | How to use UpdateLinkedDataSource with multiple worksheets | Refresh all linked objects in an Excel file using Aspose.Cells | Automate linked shape update across many sheets
// Developer Intent: Replace the external data sources of a workbook, refresh every linked shape on all worksheets, and save the updated file programmatically.
// Use Cases: Migrate a financial model to new source files and automatically update all dropdown lists, charts, and pivot tables. | Consolidate reporting data from several workbooks and refresh linked shapes across dozens of sheets in one operation. | Create a nightly job that swaps old data files with fresh extracts, ensuring all linked objects reflect the latest values.
// AI Prompts: Write C# code that loads a main workbook, swaps its external links with new workbooks, calls Shapes.UpdateSelectedValue on each worksheet, and saves the updated file using Aspose.Cells. | Explain the relationship between UpdateLinkedDataSource and Shapes.UpdateSelectedValue and why both are required for a complete refresh. | Suggest robust error‑handling patterns for missing or corrupted external workbooks when performing a batch linked‑shape update.

using System;
using System.IO;
using Aspose.Cells;

namespace LinkedShapeBatchUpdateDemo
{
    // Loads a primary workbook, replaces its external links with new source workbooks using UpdateLinkedDataSource, refreshes every linked shape (dropdowns, list boxes, etc.) on all worksheets via Shapes.UpdateSelectedValue, optionally updates dependent pivot tables and charts with RefreshAll, and saves the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Load the main workbook that contains linked shapes and external links
                const string mainPath = "MainWorkbook.xlsx";
                if (!File.Exists(mainPath))
                {
                    Console.WriteLine($"Main workbook not found: {mainPath}");
                    return;
                }
                Workbook mainWorkbook = new Workbook(mainPath);

                // Prepare external workbooks that provide the new data sources
                string[] externalPaths = { "ExternalData1.xlsx", "ExternalData2.xlsx" };
                var externalWorkbooks = new System.Collections.Generic.List<Workbook>();

                foreach (string path in externalPaths)
                {
                    if (File.Exists(path))
                    {
                        externalWorkbooks.Add(new Workbook(path));
                    }
                    else
                    {
                        Console.WriteLine($"External workbook not found and will be skipped: {path}");
                    }
                }

                // Update all external links in the main workbook with the new data sources
                if (externalWorkbooks.Count > 0)
                {
                    mainWorkbook.UpdateLinkedDataSource(externalWorkbooks.ToArray());
                }

                // Refresh all linked shapes (e.g., dropdowns, list boxes) in every worksheet
                foreach (Worksheet sheet in mainWorkbook.Worksheets)
                {
                    sheet.Shapes.UpdateSelectedValue();
                }

                // Optionally refresh all pivot tables and charts that depend on the updated data
                mainWorkbook.Worksheets.RefreshAll();

                // Save the updated workbook
                const string outputPath = "MainWorkbook_Updated.xlsx";
                mainWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
