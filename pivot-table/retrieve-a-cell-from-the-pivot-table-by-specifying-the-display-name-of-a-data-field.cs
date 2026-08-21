// Title: C# Aspose.Cells – Retrieve a Pivot Table Cell Using Data Field Display Name
// Description: This example creates a workbook, adds sample data, builds a pivot table with a row field (Fruit) and a data field (Quantity), refreshes the pivot, then uses `PivotTable.GetCellByDisplayName` to fetch the cell that holds the first data field’s display name (e.g., "Sum of Quantity"). The code prints the cell address and value before saving the file.
// Keywords: Aspose.Cells GetCellByDisplayName | pivot table cell by display name C# | Aspose.Cells pivot API | retrieve pivot data field cell .NET | C# Aspose.Cells example
// Common Searches: Aspose.Cells GetCellByDisplayName C# example | how to get pivot table cell by data field name Aspose | retrieve pivot table value using display name Aspose.Cells | C# get cell address of pivot data field Aspose
// Developer Intent: Find the worksheet cell that corresponds to a specific pivot table data field by its display name.
// Use Cases: Extract the summed value for a particular row (e.g., Fruit) to feed a report. | Apply custom formatting to the cell that contains a calculated measure such as "Sum of Quantity". | Validate pivot table results in automated tests by comparing the retrieved cell value with expected data.
// AI Prompts: Generate C# code with Aspose.Cells that locates a pivot table cell using the data field's display name and prints its address and value. | Explain the behavior of `PivotTable.GetCellByDisplayName` when the provided display name does not exist in the pivot table. | Show how to loop through all data fields in a pivot table and retrieve each associated cell with `GetCellByDisplayName`.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // This example creates a workbook, adds sample data, builds a pivot table with a row field (Fruit) and a data field (Quantity), refreshes the pivot, then uses `PivotTable.GetCellByDisplayName` to fetch the cell that holds the first data field’s display name (e.g., "Sum of Quantity"). The code prints the cell address and value before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["B3"].Value = 20;
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = 15;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add row field (Fruit) and data field (Quantity)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Refresh and calculate the pivot table so that data is populated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the display name of the first data field (e.g., "Sum of Quantity")
            string displayName = pivotTable.DataFields[0].DisplayName;

            // Use GetCellByDisplayName to obtain the cell that holds the data field value
            Cell targetCell = pivotTable.GetCellByDisplayName(displayName);

            // Output information about the retrieved cell
            Console.WriteLine("Display Name: " + displayName);
            if (targetCell != null)
            {
                Console.WriteLine("Cell Address: " + targetCell.Name);
                Console.WriteLine("Cell Value: " + targetCell.Value);
            }
            else
            {
                Console.WriteLine("Cell not found for the given display name.");
            }

            // Save the workbook to a file
            workbook.Save("PivotTable_GetCellByDisplayName_Output.xlsx");
        }
    }
}
