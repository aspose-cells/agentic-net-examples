// Title: Aspose.Cells C# – Retrieve a Pivot Table Cell Using Data Field Display Name
// Description: Demonstrates how to create a workbook, add a pivot table, obtain the display name of a data field, and use PivotTable.GetCellByDisplayName to fetch the corresponding worksheet cell. The example prints the cell address and value, then saves the file.
// Keywords: Aspose.Cells GetCellByDisplayName | C# pivot table cell by display name | Aspose.Cells PivotTable API | retrieve pivot data field cell | Aspose.Cells .NET example
// Common Searches: Aspose.Cells GetCellByDisplayName C# example | how to get pivot table cell by data field name Aspose | C# retrieve pivot table cell using display name | Aspose.Cells pivot table data field lookup
// Developer Intent: Find the worksheet cell that corresponds to a specific data field’s display name in an Aspose.Cells pivot table.
// Use Cases: Validate aggregated totals for a particular data field after pivot calculation. | Apply conditional formatting to a data‑field total cell before exporting. | Programmatically compare the same data field across multiple pivot tables.
// AI Prompts: Generate C# code with Aspose.Cells that retrieves a pivot table cell by its data field display name and handles a missing cell gracefully. | Explain what PivotTable.GetCellByDisplayName returns when the specified display name does not exist. | Create a C# snippet that loops through all data fields in a pivot table, printing each field’s display name, cell address, and value.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // Demonstrates how to create a workbook, add a pivot table, obtain the display name of a data field, and use PivotTable.GetCellByDisplayName to fetch the corresponding worksheet cell. The example prints the cell address and value, then saves the file.
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
            sheet.Cells["B3"].Value = 20;
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = 15;

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add row field and data field to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Refresh and calculate the pivot table to populate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the display name of the first data field
            string displayName = pivotTable.DataFields[0].DisplayName;

            // Use GetCellByDisplayName to obtain the corresponding cell
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

            // Save the workbook
            workbook.Save("PivotTable_GetCellByDisplayName_Output.xlsx");
        }
    }
}
