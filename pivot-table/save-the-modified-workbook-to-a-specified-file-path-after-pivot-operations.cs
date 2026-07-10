using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSaveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Define the output file path
            string outputPath = "ModifiedPivotWorkbook.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Beverage");
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["A4"].PutValue("Electronics");
            sheet.Cells["B4"].PutValue(1500);

            // Add a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, data = Sum of Amount
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Calculate data for the newly created pivot table
            pivotTable.CalculateData();

            // Refresh all pivot tables in the worksheet (good practice after modifications)
            sheet.RefreshPivotTables();

            // Optionally, control whether pivot data is saved with the workbook
            // pivotTable.SaveData = true; // default is true

            // Save the modified workbook to the specified file path
            workbook.Save(outputPath);

            // Clean up resources
            workbook.Dispose();

            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
    }
}