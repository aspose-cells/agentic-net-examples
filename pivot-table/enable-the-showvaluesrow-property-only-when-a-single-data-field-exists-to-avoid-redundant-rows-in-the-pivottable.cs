using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableShowValuesRowDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Product";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = "Apple";
            sheet.Cells["C2"].Value = 120;

            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["B3"].Value = "Orange";
            sheet.Cells["C3"].Value = 150;

            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B4"].Value = "Carrot";
            sheet.Cells["C4"].Value = 80;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields: Category as row, Sales as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Enable ShowValuesRow only when there is a single data field
            pivotTable.ShowValuesRow = pivotTable.DataFields.Count == 1;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableShowValuesRowDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}