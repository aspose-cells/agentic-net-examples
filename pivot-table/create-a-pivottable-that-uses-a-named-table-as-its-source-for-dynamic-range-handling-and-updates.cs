using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsPivotWithNamedTable
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source data on the first worksheet
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Sample data: Category, Product, Sales
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Sales");

                dataSheet.Cells["A2"].PutValue("Fruit");
                dataSheet.Cells["B2"].PutValue("Apple");
                dataSheet.Cells["C2"].PutValue(1200);

                dataSheet.Cells["A3"].PutValue("Fruit");
                dataSheet.Cells["B3"].PutValue("Banana");
                dataSheet.Cells["C3"].PutValue(800);

                dataSheet.Cells["A4"].PutValue("Vegetable");
                dataSheet.Cells["B4"].PutValue("Carrot");
                dataSheet.Cells["C4"].PutValue(600);

                dataSheet.Cells["A5"].PutValue("Vegetable");
                dataSheet.Cells["B5"].PutValue("Tomato");
                dataSheet.Cells["C5"].PutValue(900);

                // -------------------------------------------------
                // 2. Convert the range into a named table (ListObject)
                // -------------------------------------------------
                string tableRange = "A1:C5";

                // Add a ListObject; the Add method returns the index of the new table
                int tableIndex = dataSheet.ListObjects.Add("SalesTable", tableRange, true);
                ListObject table = dataSheet.ListObjects[tableIndex];
                table.TableStyleType = TableStyleType.TableStyleMedium9; // optional style

                // -------------------------------------------------
                // 3. Create a new worksheet for the PivotTable
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // -------------------------------------------------
                // 4. Add a PivotTable that uses the named table as its source
                // -------------------------------------------------
                string sourceData = "SalesTable";
                string destCell = "A1";
                string pivotName = "SalesPivot";

                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, destCell, pivotName);
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // -------------------------------------------------
                // 5. Configure the PivotTable fields
                // -------------------------------------------------
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;

                // -------------------------------------------------
                // 6. Refresh and calculate the PivotTable data
                // -------------------------------------------------
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // 7. Save the workbook
                // -------------------------------------------------
                string outputPath = "PivotTableWithNamedTable.xlsx";

                // Ensure the directory exists if a path is provided
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}