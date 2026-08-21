// Title: Create and Refresh a Pivot Table with Smart Markers Using Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to generate a workbook, insert smart‑marker data into a "SourceData" sheet, build a pivot table on a separate "PivotReport" sheet, configure row, column and data fields, refresh the pivot to capture the merged data, and save the file as PivotWithSmartMarkers.xlsx.
// Keywords: Aspose.Cells C# pivot table | smart markers Aspose.Cells | refresh pivot tables .NET | programmatic pivot report | dynamic source range Aspose | C# Excel automation | global Aspose.Cells examples
// Common Searches: how to add a pivot table with Aspose.Cells in C# | refresh pivot after data merge Aspose.Cells | use smart markers to populate Excel data for pivot | Aspose.Cells create pivot from list of objects | C# code for dynamic pivot table range
// Developer Intent: Programmatically fill a worksheet with data via smart markers, create a pivot table based on that data, refresh the pivot, and export the workbook.
// Use Cases: Generate a sales summary pivot that groups amounts by category and product without manual Excel interaction. | Automate reporting pipelines where source data changes frequently and pivots must stay up‑to‑date. | Create reusable .NET components that build Excel workbooks with smart‑marker driven data and built‑in pivot analysis.
// AI Prompts: Write C# code that uses Aspose.Cells smart markers to insert a list of objects into a worksheet and then creates a pivot table that is refreshed automatically. | Show how to determine the source range for a pivot table with MaxDisplayRange and configure row, column, and data fields in Aspose.Cells. | Explain the steps to refresh all pivot tables in an Aspose.Cells workbook after updating the source data programmatically.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSmartMarkerDemo
{
    // Sample data class for pivot processing
    // This example demonstrates how to generate a workbook, insert smart‑marker data into a "SourceData" sheet, build a pivot table on a separate "PivotReport" sheet, configure row, column and data fields, refresh the pivot to capture the merged data, and save the file as PivotWithSmartMarkers.xlsx.
    public class SalesRecord
    {
        public string Category { get; set; } = string.Empty;
        public string Product { get; set; } = string.Empty;
        public double Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Prepare the source worksheet with headers
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Amount");

                // 3. Define sample data
                List<SalesRecord> sales = new List<SalesRecord>
                {
                    new SalesRecord { Category = "Fruits", Product = "Apple",  Amount = 1200 },
                    new SalesRecord { Category = "Fruits", Product = "Banana", Amount = 800 },
                    new SalesRecord { Category = "Vegetables", Product = "Carrot", Amount = 600 },
                    new SalesRecord { Category = "Vegetables", Product = "Tomato", Amount = 950 }
                };

                // 4. Populate the source sheet with the sample data
                int currentRow = 2; // Data starts from row 2
                foreach (var record in sales)
                {
                    sourceSheet.Cells[currentRow, 0].PutValue(record.Category);
                    sourceSheet.Cells[currentRow, 1].PutValue(record.Product);
                    sourceSheet.Cells[currentRow, 2].PutValue(record.Amount);
                    currentRow++;
                }

                // 5. Create a worksheet that will hold the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotReport");

                // 6. Add a pivot table based on the populated data range
                var usedRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{usedRange.Address}";
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A3", "SalesPivot");

                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // 7. Refresh the pivot table to reflect the data
                pivotSheet.RefreshPivotTables();

                // 8. Save the workbook
                workbook.Save("PivotWithSmartMarkers.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
