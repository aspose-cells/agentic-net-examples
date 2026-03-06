using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["B3"].Value = 15;
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = 20;

            // Add a pivot table based on the data range A1:B4, place it at C3, and name it "PivotTable1"
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Fruit as row field, Quantity as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Refresh and calculate the pivot table to generate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the display name of the first data field (e.g., "Sum of Quantity")
            string displayName = pivotTable.DataFields[0].DisplayName;

            // Use GetCellByDisplayName to obtain the cell that holds the data field's total value
            Cell targetCell = pivotTable.GetCellByDisplayName(displayName);

            // Output information about the retrieved cell
            Console.WriteLine($"Display Name: {displayName}");
            if (targetCell != null)
            {
                Console.WriteLine($"Cell Name: {targetCell.Name}");
                Console.WriteLine($"Cell Value: {targetCell.Value}");
            }
            else
            {
                Console.WriteLine("Cell not found.");
            }

            // Save the workbook in XLSX format
            workbook.Save("PivotTable_GetCellByDisplayName_Output.xlsx");
        }
    }
}