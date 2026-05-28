using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotFieldListHideDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(150);

                // Add a new worksheet to host the pivot table
                int pivotSheetIndex = workbook.Worksheets.Add(SheetType.Worksheet);
                Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
                pivotSheet.Name = "PivotTable";

                // Create the pivot table (correct overload: name, sourceData, destCellName)
                int pivotTableIndex = pivotSheet.PivotTables.Add(
                    "PivotTable",   // name of the pivot table
                    "A1:B4",        // source data range
                    "C3");          // destination cell
                PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Value as data field

                // Hide the field list pane for this pivot table
                pivotTable.EnableFieldList = false;

                // Save the workbook
                workbook.Save("PivotTable_FieldListHidden.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}