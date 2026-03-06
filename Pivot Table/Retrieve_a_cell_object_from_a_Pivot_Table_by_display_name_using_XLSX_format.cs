using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main()
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

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Configure the pivot table: add row field and data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Refresh and calculate the pivot table so that data is populated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Obtain the display name of the first data field (e.g., "Sum of Quantity")
            string displayName = pivotTable.DataFields[0].DisplayName;

            // Retrieve the cell that corresponds to the display name
            Cell targetCell = pivotTable.GetCellByDisplayName(displayName);

            // Output information about the retrieved cell
            Console.WriteLine("Display Name: " + displayName);
            if (targetCell != null)
            {
                Console.WriteLine("Cell Name: " + targetCell.Name);
                Console.WriteLine("Cell Value: " + targetCell.Value);
            }
            else
            {
                Console.WriteLine("Cell not found for the given display name.");
            }

            // Save the workbook in XLSX format
            workbook.Save("PivotTable_GetCellByDisplayName_Output.xlsx");
        }
    }
}