// Title: Refresh Parent Pivot Table and Recalculate Child Pivot Table in C# with Aspose.Cells
// Description: C# example that creates a workbook, adds a parent and a child pivot table sharing the same source range, refreshes and calculates the parent first, then updates the child, and saves the file. The sequence guarantees that child totals reflect the latest parent calculations.
// Keywords: Aspose.Cells C# pivot table refresh | RefreshData CalculateData Aspose.Cells | parent child pivot tables .NET | synchronize multiple pivots Excel | pivot table totals accuracy | sample code GitHub Aspose.Cells | Excel automation C# | hierarchical pivot example
// Common Searches: how to refresh a child pivot table after parent in Aspose.Cells | C# calculate child pivot totals after parent refresh | Aspose.Cells update multiple pivot tables same source | refreshdata and calculatedata order Aspose.Cells | example code for parent‑child pivot tables .NET
// Developer Intent: Refresh the parent pivot table, then recalculate the child pivot table so that aggregated values are consistent across both tables.
// Use Cases: Generating Excel reports with hierarchical pivots where child summaries must follow parent updates. | Automating workbooks that contain several pivot tables sharing a data range and requiring synchronized calculations. | Building dashboards that reflect real‑time changes by refreshing the primary pivot and propagating results to dependent pivots.
// AI Prompts: Provide C# code using Aspose.Cells to refresh a parent pivot table and then recalculate a child pivot table for accurate totals. | Explain the required steps to keep multiple pivot tables in sync when they share the same source range in Aspose.Cells for .NET. | Suggest best practices for error handling when calling RefreshData and CalculateData on parent and child pivots with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // C# example that creates a workbook, adds a parent and a child pivot table sharing the same source range, refreshes and calculates the parent first, then updates the child, and saves the file. The sequence guarantees that child totals reflect the latest parent calculations.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot tables
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Apple");
                sheet.Cells["B4"].PutValue(80);
                sheet.Cells["A5"].PutValue("Banana");
                sheet.Cells["B5"].PutValue(70);

                // -------------------------------------------------
                // Create the parent pivot table
                // -------------------------------------------------
                PivotTableCollection pivots = sheet.PivotTables;
                int parentIndex = pivots.Add("A1:B5", "D3", "ParentPivot");
                PivotTable parentPivot = pivots[parentIndex];
                parentPivot.AddFieldToArea(PivotFieldType.Row, "Product");
                parentPivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // -------------------------------------------------
                // Create a child pivot table that uses the same data source
                // -------------------------------------------------
                int childIndex = pivots.Add("A1:B5", "D10", "ChildPivot");
                PivotTable childPivot = pivots[childIndex];
                childPivot.AddFieldToArea(PivotFieldType.Row, "Product");
                childPivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // -------------------------------------------------
                // Refresh and calculate the parent pivot table first
                // -------------------------------------------------
                parentPivot.RefreshData();      // Gather data from source to pivot cache
                parentPivot.CalculateData();    // Populate the pivot range with calculated values

                // -------------------------------------------------
                // Refresh and calculate the child pivot table
                // -------------------------------------------------
                childPivot.RefreshData();       // Refresh child's cache
                childPivot.CalculateData();     // Recalculate displayed data

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("PivotChildrenUpdated.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
