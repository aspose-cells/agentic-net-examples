using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotFieldReorderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Region";
            cells["B1"].Value = "Product";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "North";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "North";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 800;

            cells["A4"].Value = "South";
            cells["B4"].Value = "Apple";
            cells["C4"].Value = 1500;

            cells["A5"].Value = "South";
            cells["B5"].Value = "Banana";
            cells["C5"].Value = 900;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add two fields to the row area: Region (index 0) and Product (index 1)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // Region
            pivotTable.AddFieldToArea(PivotFieldType.Row, 1); // Product

            // Add the Sales field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

            // Display the original order of row fields
            Console.WriteLine("Original Row Field Order:");
            for (int i = 0; i < pivotTable.RowFields.Count; i++)
            {
                Console.WriteLine($"Position {i}: {pivotTable.RowFields[i].Name}");
            }

            // Reorder the row fields: move the field at position 0 (Region) to position 1
            // This will make Product appear before Region in the hierarchy
            pivotTable.RowFields.Move(0, 1);

            // Display the new order after moving
            Console.WriteLine("\nRow Field Order After Move:");
            for (int i = 0; i < pivotTable.RowFields.Count; i++)
            {
                Console.WriteLine($"Position {i}: {pivotTable.RowFields[i].Name}");
            }

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotFieldReorderResult.xlsx");
        }
    }
}