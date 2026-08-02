// Title: Set Absolute Row Pivot Item Position with Aspose.Cells for .NET (C#)
// Description: Shows how to create or load a workbook, add a pivot table with Category and SubCategory row fields, and use the PivotItem.PositionInSameParentNode property to place specific SubCategory items (e.g., "Beverages" first, "Snacks" second) at exact positions before saving the file.
// Keywords: Aspose.Cells | C# pivot table | PivotItem PositionInSameParentNode | absolute pivot item order | reorder row items | programmatic pivot item positioning | pivot table hierarchy | set pivot row index | .NET Excel pivot
// Common Searches: Aspose.Cells change pivot row item order | C# set pivot item position programmatically | PositionInSameParentNode usage example | how to reorder subcategory items in Aspose pivot table | move pivot item to first position with Aspose.Cells
// Developer Intent: Assign a fixed order to row items in a pivot table programmatically.
// Use Cases: Force the SubCategory "Beverages" to appear as the first row entry regardless of source data sorting. | Place the SubCategory "Snacks" in the second row position to achieve a custom display sequence. | Apply a predefined ordering to multiple row items after refreshing the pivot table for consistent reporting layouts. | Implement dynamic reordering of row items based on user‑defined priority lists.
// AI Prompts: Generate C# code that moves a pivot item named "Desserts" to the third position within its parent row field using Aspose.Cells. | Explain the PivotItem.PositionInSameParentNode property for nested row fields and show how to handle missing items gracefully. | Create a method that accepts a dictionary of pivot item names and target positions, then applies the ordering to a pivot table's row items with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotPositionDemo
{
    // Shows how to create or load a workbook, add a pivot table with Category and SubCategory row fields, and use the PivotItem.PositionInSameParentNode property to place specific SubCategory items (e.g., "Beverages" first, "Snacks" second) at exact positions before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourceFile = "SourceData.xlsx";

                // Ensure the source workbook exists; create a sample if it does not.
                if (!File.Exists(sourceFile))
                {
                    CreateSampleSourceWorkbook(sourceFile);
                }

                // Load the workbook containing source data.
                Workbook workbook = new Workbook(sourceFile);

                // Add a worksheet to host the pivot table.
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Reference the worksheet that holds the source data.
                Worksheet dataSheet = workbook.Worksheets["Data"];
                if (dataSheet == null)
                {
                    throw new InvalidOperationException("Source worksheet 'Data' not found.");
                }

                // Define the data range for the pivot table.
                string dataRange = $"='{dataSheet.Name}'!A1:D100";

                // Add the pivot table to the pivot sheet at cell A3.
                int ptIndex = pivotSheet.PivotTables.Add(dataRange, "A3", "MyPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[ptIndex];

                // Add row fields (Category -> SubCategory) and a data field.
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Populate the pivot table.
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Access the innermost row field's items.
                PivotItemCollection subItems = pivotTable.RowFields["SubCategory"].PivotItems;

                // Move specific items to desired absolute positions.
                if (subItems["Beverages"] != null)
                {
                    subItems["Beverages"].PositionInSameParentNode = 0; // first position
                }

                if (subItems["Snacks"] != null)
                {
                    subItems["Snacks"].PositionInSameParentNode = 1; // second position
                }

                // Recalculate after position changes.
                pivotTable.CalculateData();

                // Save the updated workbook.
                workbook.Save("PivotWithAbsolutePositions.xlsx");
                Console.WriteLine("Pivot table created and saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a simple source workbook with sample data if the file is missing.
        private static void CreateSampleSourceWorkbook(string filePath)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Header row
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("SubCategory");
            ws.Cells["C1"].PutValue("Product");
            ws.Cells["D1"].PutValue("Amount");

            // Sample data
            string[,] data = new string[,]
            {
                { "Beverages", "Tea", "Green Tea", "120" },
                { "Beverages", "Coffee", "Espresso", "200" },
                { "Snacks", "Chips", "Potato Chips", "80" },
                { "Snacks", "Nuts", "Almonds", "150" }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                ws.Cells[i + 1, 0].PutValue(data[i, 0]);
                ws.Cells[i + 1, 1].PutValue(data[i, 1]);
                ws.Cells[i + 1, 2].PutValue(data[i, 2]);
                ws.Cells[i + 1, 3].PutValue(Convert.ToDouble(data[i, 3]));
            }

            wb.Save(filePath);
        }
    }
}
