using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableDestinationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Fruits";
            sheet.Cells["B2"].Value = 1500;
            sheet.Cells["A3"].Value = "Vegetables";
            sheet.Cells["B3"].Value = 2500;
            sheet.Cells["A4"].Value = "Fruits";
            sheet.Cells["B4"].Value = 1200;
            sheet.Cells["A5"].Value = "Vegetables";
            sheet.Cells["B5"].Value = 1800;

            // Define the source data range (including headers)
            string sourceData = "A1:B5";

            // Define the destination cell where the pivot table will be placed
            string destCell = "D3";

            // Define a name for the pivot table
            string tableName = "SalesPivot";

            // Add the pivot table
            int pivotIndex = sheet.PivotTables.Add(sourceData, destCell, tableName);
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Calculate the pivot data
            pivot.CalculateData();

            // Save the workbook
            string filePath = "PivotTableWithDestination.xlsx";
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            workbook.Save(filePath);
        }
    }
}