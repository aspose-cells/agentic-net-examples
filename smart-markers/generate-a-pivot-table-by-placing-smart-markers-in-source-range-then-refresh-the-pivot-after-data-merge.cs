// Title: Generate a pivot table from smart markers and refresh it after data merge using Aspose.Cells for .NET (C#)
// AI Prompts: Build a workbook that inserts `${Product}` and `${Sales}` smart‑marker placeholders, fills them from a DataTable, creates a pivot table on the expanded range, refreshes the pivot, and saves the file. | Extend the example to add a second data field (e.g., Quantity) to the pivot table, recalculate the pivot after WorkbookDesigner processes the smart markers, and export the updated workbook. | Demonstrate how to loop through a worksheet's PivotTables collection and invoke Refresh after calling WorkbookDesigner.Process() to merge smart‑marker data.
// Common Searches: Aspose.Cells C# create pivot table after smart marker processing | how to refresh worksheet pivot tables after WorkbookDesigner merges data | using smart markers to populate pivot table source range in .NET
// Tags: Aspose.Cells smart markers pivot table generation | C# refresh pivot tables after WorkbookDesigner | dynamic pivot source range with smart markers | WorkbookDesigner process DataTable example | calculate pivot data programmatically Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSmartMarkerDemo
{
    // The sample creates a new workbook, places `${Product}` and `${Sales}` smart‑marker placeholders, populates them from a DataTable via WorkbookDesigner, defines a pivot table on the resulting range, refreshes and calculates the pivot, and saves the workbook as PivotTable_With_SmartMarkers.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (will hold source data and smart markers)
                Worksheet dataSheet = workbook.Worksheets[0];
                Cells cells = dataSheet.Cells;

                // Add headers
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Sales");

                // Insert smart markers – these will be replaced by actual data during processing
                cells["A2"].PutValue("${Product}");
                cells["B2"].PutValue("${Sales}");

                // Create a pivot table that will use the data range after smart marker processing.
                // Initially point it to the header row; the range will expand automatically.
                int pivotIndex = dataSheet.PivotTables.Add("A1:B2", "D4", "SalesPivot");
                PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Prepare sample data for smart marker processing
                DataTable dt = new DataTable();
                dt.Columns.Add("Product", typeof(string));
                dt.Columns.Add("Sales", typeof(double));

                dt.Rows.Add("Apple", 1200);
                dt.Rows.Add("Banana", 850);
                dt.Rows.Add("Orange", 430);
                dt.Rows.Add("Grape", 670);

                // Process smart markers using WorkbookDesigner (correct API)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // After data is merged, refresh the pivot table to reflect the new source data
                dataSheet.RefreshPivotTables();

                // Optionally calculate the pivot data (ensures values are written to cells)
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_With_SmartMarkers.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
