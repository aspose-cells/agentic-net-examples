using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableHideAndSortDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Food");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Beverage");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Beverage");
            dataSheet.Cells["B5"].PutValue(200);
            dataSheet.Cells["A6"].PutValue("Electronics");
            dataSheet.Cells["B6"].PutValue(500);
            dataSheet.Cells["A7"].PutValue("Electronics");
            dataSheet.Cells["B7"].PutValue(300);

            // Add a new worksheet to host the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add(SheetType.Worksheet);
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotTable";

            // Create the pivot table (source range A1:B7, destination start cell C3)
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B7", "C3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Category as row, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Hide the pivot field list UI
            workbook.Settings.HidePivotFieldList = true;

            // Hide all categories except "Food"
            PivotField rowField = pivotTable.RowFields[0];
            for (int i = 0; i < rowField.ItemCount; i++)
            {
                string item = rowField.Items[i];
                bool hide = !item.Equals("Food", StringComparison.OrdinalIgnoreCase);
                rowField.HideItem(i, hide);
            }

            // Enable custom list sort (allows Excel to use its built‑in custom list ordering)
            pivotTable.CustomListSort = true;

            // Calculate the pivot data so that changes are reflected
            pivotTable.CalculateData();

            // Refresh all pivot tables in the workbook (good practice after modifications)
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook as XLSX
            workbook.Save("PivotTableHideAndSortDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableHideAndSortDemo.Run();
        }
    }
}