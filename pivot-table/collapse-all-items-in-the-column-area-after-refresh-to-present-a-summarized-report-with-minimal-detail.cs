// Title: Collapse All Column Items in an Aspose.Cells Pivot Table After Refresh (C#)
// Description: Demonstrates how to refresh a pivot table, then programmatically hide the details of every column field so the report shows only aggregated totals, producing a concise summary view.
// Keywords: Aspose.Cells collapse pivot columns | C# hide pivot column details | refresh pivot table Aspose.Cells | PivotField HideDetail example | Aspose.Cells pivot drill‑down collapse | summarized pivot report C#
// Common Searches: collapse column fields after RefreshData Aspose.Cells | C# hide detail for all pivot column items | programmatically collapse pivot table columns .NET | Aspose.Cells pivot table summary view | how to hide column details in Aspose pivot
// Developer Intent: Programmatically collapse every column‑area item in a pivot table after refreshing its data.
// Use Cases: Create a sales dashboard that displays only total sales per product without expanding each product column. | Build a workbook template that automatically presents a clean, aggregated view after data refresh. | Prepare a distribution‑ready report where end users see only high‑level column headings.
// AI Prompts: Generate C# code using Aspose.Cells to refresh a pivot table and collapse all column field items. | Show how to enable drill‑down, refresh data, and call HideDetail(true) on each column field in a pivot table. | Explain the effect of PivotField.HideDetail(true) on column areas and how to apply it across all columns.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to refresh a pivot table, then programmatically hide the details of every column field so the report shows only aggregated totals, producing a concise summary view.
    public class CollapseColumnAreaAfterRefresh
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Region");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue("North");
                sheet.Cells["B2"].PutValue("Widget");
                sheet.Cells["C2"].PutValue(1200);

                sheet.Cells["A3"].PutValue("North");
                sheet.Cells["B3"].PutValue("Gadget");
                sheet.Cells["C3"].PutValue(800);

                sheet.Cells["A4"].PutValue("South");
                sheet.Cells["B4"].PutValue("Widget");
                sheet.Cells["C4"].PutValue(1500);

                sheet.Cells["A5"].PutValue("South");
                sheet.Cells["B5"].PutValue("Gadget");
                sheet.Cells["C5"].PutValue(700);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");          // Row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");     // Column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");         // Data field

                // Enable drill-down and show expand/collapse buttons
                pivotTable.EnableDrilldown = true;
                pivotTable.ShowDrill = true;

                // Refresh the pivot cache and calculate data
                pivotTable.RefreshData();      // Correct API call
                pivotTable.CalculateData();

                // Collapse all items in the column area
                foreach (PivotField columnField in pivotTable.ColumnFields)
                {
                    columnField.HideDetail(true); // Hide detail (collapse)
                }

                // Save the workbook with the collapsed column area
                string outputPath = "CollapsedColumnArea.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the example runner
    public class Program
    {
        public static void Main(string[] args)
        {
            CollapseColumnAreaAfterRefresh.Run();
        }
    }
}
