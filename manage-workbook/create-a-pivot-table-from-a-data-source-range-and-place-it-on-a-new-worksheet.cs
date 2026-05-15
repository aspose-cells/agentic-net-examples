using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample source data
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Sample data: Category and Value columns
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["A4"].PutValue("A");
            sourceSheet.Cells["B4"].PutValue(30);

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Determine the source data range (including all populated cells)
            AsposeRange sourceRange = sourceSheet.Cells.MaxDisplayRange;
            // Build the source data string in the required format
            string sourceData = $"=SourceData!{sourceRange.Address}";

            // Add the pivot table to the new worksheet using the Add(string, string, string) method
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, "A1", "MyPivotTable");

            // Configure the pivot table (e.g., Category as row field, Value as data field)
            PivotTable pivotTable = pivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Save the workbook to a file
            workbook.Save("PivotTableDemo.xlsx");
        }
    }
}