using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsGroupExample
{
    public class Program
    {
        public static void Main()
        {
            // Load options
            LoadOptions loadOptions = new LoadOptions { IgnoreUselessShapes = true };
            Workbook workbook = new Workbook("input.xlsx", loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Group rows (rows 2‑4) and hide them
            cells.GroupRows(1, 3, true);
            cells.ShowGroupDetail(false, 1);

            // Group columns (B‑D) and hide them
            cells.GroupColumns(1, 3, true);
            cells.ShowGroupDetail(true, 1);

            // Sample data for pivot table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("A");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("B");
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["A5"].PutValue("B");
            worksheet.Cells["B5"].PutValue(40);

            // Add pivot table
            int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D1", "SamplePivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            int valueRowFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "Value");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Group the numeric row field by interval 15 (0‑50)
            PivotField valueRowField = pivotTable.RowFields[valueRowFieldIndex];
            valueRowField.GroupBy(0, 50, 15, false);

            // Refresh and calculate pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save workbook
            workbook.Save("output.xlsx");
        }
    }
}