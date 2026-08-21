// Title: C# – Load Excel Workbook without Pivot Cache and Refresh Pivot Tables using Aspose.Cells
// Description: Demonstrates how to create LoadOptions with ParsingPivotCachedRecords set to false, load an XLSX workbook, refresh all pivot tables to reflect current data, and save the result. Skipping cached pivot records speeds up loading large files before updating the pivots.
// Keywords: Aspose.Cells LoadOptions ParsingPivotCachedRecords false | disable pivot cache parsing .NET | refresh all pivot tables Aspose.Cells | C# Excel pivot table refresh | improve workbook load performance | Aspose.Cells performance tuning | Excel pivot cache skip
// Common Searches: Aspose.Cells load workbook without pivot cache | How to refresh pivot tables after loading Excel with Aspose.Cells | ParsingPivotCachedRecords property example C# | Speed up loading large Excel files Aspose.Cells | Refresh pivot tables programmatically .NET
// Developer Intent: Load an Excel file without parsing cached pivot data and then refresh its pivot tables.
// Use Cases: Accelerate loading of massive workbooks by ignoring stored pivot caches, then generate up‑to‑date reports. | Automate post‑processing of template files that contain pivots after source data changes. | Integrate fast workbook loading into .NET reporting pipelines where pivot tables must reflect the latest data.
// AI Prompts: Generate C# code that opens an XLSX file with LoadOptions.ParsingPivotCachedRecords set to false, refreshes all pivot tables, and saves the workbook using Aspose.Cells. | Explain why disabling ParsingPivotCachedRecords improves load time and outline the steps required to correctly refresh pivot tables afterward.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    // Demonstrates how to create LoadOptions with ParsingPivotCachedRecords set to false, load an XLSX workbook, refresh all pivot tables to reflect current data, and save the result. Skipping cached pivot records speeds up loading large files before updating the pivots.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that contains a pivot table
            string sourcePath = "input.xlsx";
            // Path where the refreshed workbook will be saved
            string outputPath = "output.xlsx";

            // Create LoadOptions and disable parsing of pivot cached records
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.ParsingPivotCachedRecords = false;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Refresh all pivot tables in the workbook
            // This ensures that the pivot tables reflect the current source data
            workbook.Worksheets.RefreshPivotTables();

            // Save the refreshed workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
