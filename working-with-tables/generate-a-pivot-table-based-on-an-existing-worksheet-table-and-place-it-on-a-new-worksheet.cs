using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsPivotExample
{
    public class CreatePivotFromTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Source worksheet with data and a table ----------
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample data
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                for (int i = 2; i <= 10; i++)
                {
                    sourceSheet.Cells[$"A{i}"].PutValue("Cat" + ((i % 3) + 1));
                    sourceSheet.Cells[$"B{i}"].PutValue("Prod" + i);
                    sourceSheet.Cells[$"C{i}"].PutValue(i * 100);
                }

                // Define a table (ListObject) over the data range
                int firstDataRow = 0; // zero‑based index
                int firstDataColumn = 0;
                int totalRows = sourceSheet.Cells.MaxDisplayRange.RowCount;
                int totalColumns = sourceSheet.Cells.MaxDisplayRange.ColumnCount;

                int tableIndex = sourceSheet.ListObjects.Add(firstDataRow, firstDataColumn,
                                                             totalRows - 1, totalColumns - 1, true);
                ListObject table = sourceSheet.ListObjects[tableIndex];
                // Set the table name (use DisplayName as Name is not available in this version)
                table.DisplayName = "SalesTable";

                // ---------- Destination worksheet for the pivot table ----------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Build the source data reference using the table name
                // For a table, the reference format is: =SheetName!TableName
                string sourceData = $"=SourceData!{table.DisplayName}";

                // Add a new pivot table to the destination sheet (cell A1, name "SalesPivot")
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product"); // Column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // Data field (sum)

                // Refresh the pivot table to calculate data
                pivotSheet.RefreshPivotTables();

                // Save the workbook
                string outputPath = "PivotFromTableDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            CreatePivotFromTable.Run();
        }
    }
}