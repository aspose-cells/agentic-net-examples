using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsPivotJsonExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet (data source)
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data for the pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Quantity");
                dataSheet.Cells["A2"].PutValue("Fruit");
                dataSheet.Cells["B2"].PutValue("Apple");
                dataSheet.Cells["C2"].PutValue(120);
                dataSheet.Cells["A3"].PutValue("Fruit");
                dataSheet.Cells["B3"].PutValue("Banana");
                dataSheet.Cells["C3"].PutValue(85);
                dataSheet.Cells["A4"].PutValue("Vegetable");
                dataSheet.Cells["B4"].PutValue("Carrot");
                dataSheet.Cells["C4"].PutValue(60);
                dataSheet.Cells["A5"].PutValue("Vegetable");
                dataSheet.Cells["B5"].PutValue("Tomato");
                dataSheet.Cells["C5"].PutValue(45);

                // Add a new worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Define the source range for the pivot table (A1:C5)
                int firstRow = 0, firstColumn = 0, totalRows = 5, totalColumns = 3;
                string startCell = CellsHelper.CellIndexToName(firstRow, firstColumn);
                string endCell = CellsHelper.CellIndexToName(firstRow + totalRows - 1, firstColumn + totalColumns - 1);
                string sourceRange = $"=Data!{startCell}:{endCell}";

                // Add the pivot table at cell A1 of the pivot sheet
                int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure fields (Category as row, Product as column, Quantity as data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Get the underlying data range (the same range used as the data source)
                Aspose.Cells.Range sourceRangeObj = dataSheet.Cells.CreateRange(startCell, endCell);

                // Configure JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,
                    ExportEmptyCells = true,
                    ExportAsString = false,
                    Indent = "  "
                };

                // Convert the range to JSON string
                string jsonResult = sourceRangeObj.ToJson(jsonOptions);

                // Output the JSON string
                Console.WriteLine("Pivot Table Underlying Data as JSON:");
                Console.WriteLine(jsonResult);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}