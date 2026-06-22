using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Source data worksheet ----------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Header
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Value");

            // Initial rows
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["A4"].PutValue("A");
            sourceSheet.Cells["B4"].PutValue(30);

            // Additional rows to demonstrate that the range will expand automatically
            sourceSheet.Cells["A5"].PutValue("C");
            sourceSheet.Cells["B5"].PutValue(40);
            sourceSheet.Cells["A6"].PutValue("B");
            sourceSheet.Cells["B6"].PutValue(50);

            // ---------- Pivot table worksheet ----------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Use MaxDisplayRange to obtain a dynamic source range that grows with the data
            AsposeRange sourceRange = sourceSheet.Cells.MaxDisplayRange;
            string sourceData = $"=SourceData!{sourceRange.Address}";

            // Add a pivot table using the dynamic source range
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, "A1", "MyPivotTable");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate to ensure the pivot reflects the current source data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("DynamicPivotTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}