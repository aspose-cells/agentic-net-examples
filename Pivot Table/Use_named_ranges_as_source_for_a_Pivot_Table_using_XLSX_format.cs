using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableWithNamedRange
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (source data)
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Populate sample data
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Amount");
            sourceSheet.Cells["A2"].PutValue("Food");
            sourceSheet.Cells["B2"].PutValue(120);
            sourceSheet.Cells["A3"].PutValue("Drink");
            sourceSheet.Cells["B3"].PutValue(80);
            sourceSheet.Cells["A4"].PutValue("Food");
            sourceSheet.Cells["B4"].PutValue(150);
            sourceSheet.Cells["A5"].PutValue("Drink");
            sourceSheet.Cells["B5"].PutValue(70);

            // Define a named range that covers the data (including headers)
            int nameIndex = workbook.Worksheets.Names.Add("MyDataRange");
            workbook.Worksheets.Names[nameIndex].RefersTo = "=SourceData!$A$1:$B$5";

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Use the named range as the source for the pivot table.
            string sourceData = "MyDataRange";
            string destCell = "A1";
            string pivotName = "SalesPivot";

            // Add the pivot table
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, destCell, pivotName);

            // Configure the pivot table
            PivotTable pivotTable = pivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivotTable.RefreshData();

            // Save the workbook
            workbook.Save("PivotTableWithNamedRange.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            PivotTableWithNamedRange.Run();
        }
    }
}