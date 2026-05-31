using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (source data)
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Populate sample data for the pivot table
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Amount");
            sourceSheet.Cells["A2"].PutValue("Food");
            sourceSheet.Cells["B2"].PutValue(120);
            sourceSheet.Cells["A3"].PutValue("Beverage");
            sourceSheet.Cells["B3"].PutValue(80);
            sourceSheet.Cells["A4"].PutValue("Food");
            sourceSheet.Cells["B4"].PutValue(150);
            sourceSheet.Cells["A5"].PutValue("Beverage");
            sourceSheet.Cells["B5"].PutValue(70);

            // Add a new worksheet that will host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Define the source data range (including sheet name)
            // Using the A1 style address with sheet reference
            string sourceData = $"=SourceData!A1:B5";

            // Destination cell where the pivot table will start (upper‑left corner)
            string destCellName = "C3";

            // Name of the new pivot table
            string tableName = "SalesPivot";

            // Add the pivot table using the (string, string, string) overload
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, destCellName, tableName);

            // Retrieve the created pivot table
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot table: put "Category" in rows and "Amount" in data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableDemo.xlsx");
        }
    }
}