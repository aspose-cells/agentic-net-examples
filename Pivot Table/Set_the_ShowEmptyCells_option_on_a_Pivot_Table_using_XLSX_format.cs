using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableShowEmptyCellsDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data (including empty cells)
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Electronics");
            dataSheet.Cells["B2"].PutValue("TV");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("Electronics");
            dataSheet.Cells["B3"].PutValue("");          // empty product
            dataSheet.Cells["C3"].PutValue("");          // empty sales

            dataSheet.Cells["A4"].PutValue("Furniture");
            dataSheet.Cells["B4"].PutValue("Chair");
            dataSheet.Cells["C4"].PutValue(300);

            dataSheet.Cells["A5"].PutValue("Furniture");
            dataSheet.Cells["B5"].PutValue("");          // empty product
            dataSheet.Cells["C5"].PutValue("");          // empty sales

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table based on the data range
            // Source range: A1:C5, Destination top‑left cell: A1, Pivot table name: PivotTable1
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A1", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Set ShowEmptyCol and ShowEmptyRow to true to display empty columns/rows
            pivotTable.ShowEmptyCol = true;   // Show empty columns
            pivotTable.ShowEmptyRow = true;   // Show empty rows

            // Refresh data and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotTableShowEmptyCells.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            PivotTableShowEmptyCellsDemo.Run();
        }
    }
}