using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableCustomSortingDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source data
            // Column A – Category, Column B – Value
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Low";
            cells["B2"].Value = 10;
            cells["A3"].Value = "Medium";
            cells["B3"].Value = 30;
            cells["A4"].Value = "High";
            cells["B4"].Value = 20;
            cells["A5"].Value = "Low";
            cells["B5"].Value = 15;
            cells["A6"].Value = "Medium";
            cells["B6"].Value = 25;
            cells["A7"].Value = "High";
            cells["B7"].Value = 35;

            // Define a custom sort order for the Category field:
            // Desired order: Medium, High, Low
            // Use DataSorter to sort the source range according to this custom list
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true; // First row contains headers
            sorter.AddKey(0, SortOrder.Ascending, "Medium,High,Low"); // Column A, custom list
            sorter.Sort(cells, CellArea.CreateCellArea("A1", "B7"));

            // Add a pivot table based on the sorted source data
            int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add row field (Category) and data field (Amount)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable built‑in custom list sorting for the pivot table
            pivotTable.CustomListSort = true;

            // Refresh and calculate the pivot table to reflect the sorted data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotTable_CustomSorting_Example.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            PivotTableCustomSortingDemo.Run();
        }
    }
}