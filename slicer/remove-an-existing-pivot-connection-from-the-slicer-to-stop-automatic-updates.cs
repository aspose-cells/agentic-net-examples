// Title: Remove Pivot Table Connection from a Slicer in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, build a pivot table, insert a slicer linked to the "Product" field, and then call slicer.RemovePivotConnection(pivotTable) to detach the slicer so it no longer updates automatically when the pivot table changes. The workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells RemovePivotConnection | C# slicer detach from pivot | stop slicer auto refresh | pivot table slicer disconnect .NET | Aspose.Cells slicer API example
// Common Searches: how to detach a slicer from a pivot table using Aspose.Cells | remove pivot connection from slicer C# | prevent slicer auto refresh Aspose.Cells .NET | Aspose.Cells slicer RemovePivotConnection usage
// Developer Intent: Detach a slicer from its associated pivot table so that subsequent pivot refreshes do not modify the slicer state.
// Use Cases: Create a reporting template where the slicer remains static while the pivot data is refreshed later. | Generate a workbook for end‑users that includes slicers for manual filtering only, avoiding automatic changes after data updates. | Export data with slicers, then remove their pivot connections before saving to prevent unintended UI changes in the final file.
// AI Prompts: Provide a C# Aspose.Cells example that builds a pivot table, adds a slicer, and then removes the pivot connection with RemovePivotConnection. | Explain the effect of RemovePivotConnection on slicer behavior and how to re‑link a slicer to a pivot table if needed. | Show step‑by‑step code to detach a slicer from its pivot table to stop automatic updates, then save the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, build a pivot table, insert a slicer linked to the "Product" field, and then call slicer.RemovePivotConnection(pivotTable) to detach the slicer so it no longer updates automatically when the pivot table changes. The workbook is saved as an .xlsx file.
    public class RemovePivotConnectionFromSlicer
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["A4"].PutValue("Apple");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["B4"].PutValue(150);

            // Add a worksheet to host the pivot table and slicer
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

            // Create a pivot table based on the data range
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add("Data!A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a slicer linked to the pivot table (field "Product")
            int slicerIndex = pivotSheet.Slicers.Add(pivotTable, "E3", "Product");
            Slicer slicer = pivotSheet.Slicers[slicerIndex];

            // Remove the pivot connection from the slicer to stop automatic updates
            slicer.RemovePivotConnection(pivotTable);

            // Save the workbook
            workbook.Save("RemovePivotConnectionDemo.xlsx");
        }
    }
}
