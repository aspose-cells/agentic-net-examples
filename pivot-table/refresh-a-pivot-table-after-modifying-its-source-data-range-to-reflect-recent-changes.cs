// Title: Refresh a PivotTable after source data changes with Aspose.Cells (C#)
// Description: Demonstrates how to modify source cells, then call RefreshData and CalculateData on an Aspose.Cells PivotTable to keep the report up‑to‑date before saving the workbook.
// Keywords: Aspose.Cells PivotTable refresh C# | RefreshData CalculateData Aspose.Cells | update pivot source range .NET | recalculate pivot after data edit | C# Excel pivot table automation
// Common Searches: how to refresh pivot table in Aspose.Cells C# | Aspose.Cells RefreshData vs CalculateData | update pivot source data programmatically | C# code to recalc Excel pivot after changes | Aspose.Cells pivot table refresh example
// Developer Intent: Synchronize a PivotTable with modified source data by invoking the appropriate refresh and calculation methods.
// Use Cases: After adjusting sales figures in code, call RefreshData and CalculateData to reflect the new totals in the pivot. | When rows are added or removed from the source range, use RefreshData to include them and then recalculate the pivot layout. | In scheduled report generation, automatically update source values and ensure all pivots are current before exporting.
// AI Prompts: Write C# code that updates several source cells and then refreshes all PivotTables in a workbook using Aspose.Cells. | Explain the difference between RefreshData and CalculateData in Aspose.Cells PivotTable handling and give examples of when each is required. | Create a utility method that scans a workbook, finds every PivotTable, and performs a refresh‑calculate cycle after data modifications.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotRefreshExample
{
    // Demonstrates how to modify source cells, then call RefreshData and CalculateData on an Aspose.Cells PivotTable to keep the report up‑to‑date before saving the workbook.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet dataSheet = workbook.Worksheets[0];           // default sheet for data

            // ---------- Populate source data ----------
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(1200);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(850);
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(950);

            // ---------- Add a worksheet for the pivot table ----------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // ---------- Create the pivot table ----------
            // Source range: A1:B4 on the data sheet, Destination: C3 on the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Product as row, Sales as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Initial calculation so the pivot shows data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ---------- Modify the source data ----------
            dataSheet.Cells["B2"].PutValue(1300);   // Apple sales increased
            dataSheet.Cells["B3"].PutValue(900);    // Banana sales increased

            // ---------- Refresh the pivot table to reflect changes ----------
            pivotTable.RefreshData();               // gather new data from source
            pivotTable.CalculateData();             // recalculate pivot results

            // ---------- Save the workbook ----------
            workbook.Save("PivotRefreshResult.xlsx"); // save
        }
    }
}
