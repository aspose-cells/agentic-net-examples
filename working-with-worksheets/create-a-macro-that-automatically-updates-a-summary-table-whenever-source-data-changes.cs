// Title: C# Aspose.Cells macro to auto‑refresh a PivotTable summary when source data changes
// Description: Creates a workbook with a source data sheet, adds a PivotTable summary on a separate sheet with ManualUpdate enabled, modifies source cells and adds rows, then programmatically calls RefreshData and CalculateData to keep the summary up‑to‑date before saving the file.
// Keywords: Aspose.Cells | C# | .NET | PivotTable refresh | ManualUpdate | auto update summary table | programmatic pivot refresh | Excel macro alternative | dynamic summary sheet | Aspose.Cells API
// Common Searches: Aspose.Cells refresh pivot after data change | C# example manual update pivot Aspose.Cells | auto‑refresh summary sheet Aspose.Cells macro | how to recalculate PivotTable programmatically Aspose.Cells | Aspose.Cells pivot table dynamic update
// Developer Intent: Programmatically refresh a PivotTable summary whenever the underlying worksheet data is edited, using Aspose.Cells in C#.
// Use Cases: Generate a sales‑by‑category report that instantly reflects inventory adjustments. | Add new product rows to a data sheet and have category totals update without manual interaction. | Build a batch‑processing workbook where bulk edits trigger automatic recalculation of all aggregated fields.
// AI Prompts: Write C# code with Aspose.Cells that creates a PivotTable, sets ManualUpdate to true, modifies source data, and then refreshes the pivot. | Provide an Aspose.Cells macro that detects changes in a worksheet and automatically calls RefreshData and CalculateData on related PivotTables. | Explain how to configure a workbook so a summary PivotTable reflects newly added rows without user intervention.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsMacroDemo
{
    // Creates a workbook with a source data sheet, adds a PivotTable summary on a separate sheet with ManualUpdate enabled, modifies source cells and adds rows, then programmatically calls RefreshData and CalculateData to keep the summary up‑to‑date before saving the file.
    public class SummaryTableUpdater
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source data on the first worksheet
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "SourceData";

                // Header row
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Item");
                dataSheet.Cells["C1"].PutValue("Quantity");
                dataSheet.Cells["D1"].PutValue("Price");

                // Sample data rows
                object[,] data = new object[,]
                {
                    {"Fruit", "Apple",  10, 1.20},
                    {"Fruit", "Banana", 15, 0.80},
                    {"Fruit", "Orange", 12, 1.00},
                    {"Veg",   "Carrot", 20, 0.50},
                    {"Veg",   "Tomato", 18, 0.70},
                    {"Veg",   "Pepper", 10, 1.10}
                };

                for (int r = 0; r < data.GetLength(0); r++)
                    for (int c = 0; c < data.GetLength(1); c++)
                        dataSheet.Cells[r + 1, c].PutValue(data[r, c]);

                // -------------------------------------------------
                // 2. Create a summary table (PivotTable) on a new sheet
                // -------------------------------------------------
                Worksheet summarySheet = workbook.Worksheets.Add("Summary");
                // Define the source range (including headers)
                string sourceRange = $"=SourceData!{dataSheet.Cells.MaxDisplayRange.Address}";
                // Add the pivot table; it will be placed starting at cell A3
                int pivotIndex = summarySheet.PivotTables.Add(sourceRange, "A3", "SalesSummary");
                PivotTable pivot = summarySheet.PivotTables[pivotIndex];

                // Configure the pivot: Category as row, Item as column, Sum of Quantity and Price as data
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Column, "Item");
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
                pivot.AddFieldToArea(PivotFieldType.Data, "Price");

                // Enable manual update so the pivot does NOT refresh automatically
                pivot.ManualUpdate = true;

                // Initial refresh to populate the summary table
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // 3. Simulate a change in the source data
                // -------------------------------------------------
                // For example, increase the quantity of Apples and add a new row
                dataSheet.Cells["C2"].PutValue(25); // Apple quantity from 10 to 25
                int newRow = dataSheet.Cells.MaxDataRow + 1;
                dataSheet.Cells[newRow, 0].PutValue("Fruit");   // Category
                dataSheet.Cells[newRow, 1].PutValue("Grapes");  // Item
                dataSheet.Cells[newRow, 2].PutValue(8);        // Quantity
                dataSheet.Cells[newRow, 3].PutValue(2.00);     // Price

                // -------------------------------------------------
                // 4. Refresh the summary table to reflect changes
                // -------------------------------------------------
                // Since ManualUpdate is true, we need to refresh explicitly
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save("SummaryTableUpdated.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SummaryTableUpdater.Run();
        }
    }
}
