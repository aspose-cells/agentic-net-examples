// Title: How to retrieve a pivot table cell by its data field display name with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, obtains the display name of a data field, and uses PivotTable.GetCellByDisplayName to fetch the corresponding cell. | Show how to call GetCellByDisplayName on a PivotTable object to locate the cell containing the calculated "Sum of Quantity" value. | Demonstrate extracting the address and value of a pivot table data field cell using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# get pivot table cell using data field display name | PivotTable.GetCellByDisplayName example in .NET | How to find the cell for "Sum of Quantity" in an Aspose.Cells pivot table | Retrieve calculated pivot table value by display name with Aspose.Cells C#
// Tags: Aspose.Cells PivotTable GetCellByDisplayName | C# retrieve pivot table data field cell | Aspose.Cells calculate pivot table values | pivot table cell address extraction Aspose.Cells | display name based cell lookup Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // This C# example creates a workbook, adds a pivot table summarizing fruit quantities, refreshes and calculates it, reads the display name of the first data field (e.g., "Sum of Quantity"), uses PivotTable.GetCellByDisplayName to locate the cell that holds that value, prints the cell address and value, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
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
            int ptIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add a row field and a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Refresh and calculate the pivot table so that values are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the display name of the first data field (e.g., "Sum of Quantity")
            string displayName = pivotTable.DataFields[0].DisplayName;

            // Use GetCellByDisplayName to obtain the cell that holds the data field's value
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

            // Save the workbook
            workbook.Save("PivotTable_GetCellByDisplayName_Output.xlsx");
        }
    }
}
