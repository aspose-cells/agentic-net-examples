using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 1200;
            sheet.Cells["A3"].Value = "Banana";
            sheet.Cells["B3"].Value = 850;
            sheet.Cells["A4"].Value = "Orange";
            sheet.Cells["B4"].Value = 950;

            // Add a pivot table covering the data range and place it at C7
            int ptIndex = sheet.PivotTables.Add("A1:B4", "C7", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add fields: Product as row, Sales as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table to populate values
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the display name of the first data field
            string displayName = pivotTable.DataFields[0].DisplayName;

            // Use GetCellByDisplayName to obtain the cell that holds the data field value
            Cell targetCell = pivotTable.GetCellByDisplayName(displayName);

            // Output information about the retrieved cell
            Console.WriteLine($"Display Name: {displayName}");
            if (targetCell != null)
            {
                Console.WriteLine($"Cell Address: {targetCell.Name}");
                Console.WriteLine($"Cell Value: {targetCell.Value}");
            }
            else
            {
                Console.WriteLine("Cell not found.");
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save("PivotTable_GetCellByDisplayName_Output.xlsx");
        }
    }
}