// Title: How to get the absolute Position of a PivotItem in a row field using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a pivot table with Aspose.Cells, accesses the first PivotItem of a row field, and reads its Position property. | Show how to output the absolute Position of a PivotItem after calling RefreshData and CalculateData on the pivot table in Aspose.Cells.
// Common Searches: Aspose.Cells C# read PivotItem.Position property | Get index of first row field item in Aspose.Cells pivot table | How to retrieve absolute position of a pivot item in .NET | C# Aspose.Cells pivot table row field item order | Access PivotItem.Position after RefreshData and CalculateData
// Tags: Aspose.Cells pivot item position | C# read PivotItem.Position | Aspose.Cells get row field item index | pivot table item order Aspose.Cells | Aspose.Cells refresh calculate pivot data

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemPositionDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, assigns the Product column as a row field, refreshes and calculates the pivot table, then retrieves the first PivotItem of that row field, reads its absolute Position property, prints the value, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(3000);

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add the "Product" column as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add the "Sales" column as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table so that PivotItems are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the first row field (Product) and its first PivotItem
            PivotField rowField = pivotTable.RowFields[0];
            PivotItem firstItem = rowField.PivotItems[0];

            // Read the absolute Position property of the PivotItem
            int absolutePosition = firstItem.Position;

            // Output the position to the console
            Console.WriteLine($"Absolute Position of PivotItem \"{firstItem.Name}\": {absolutePosition}");

            // Save the workbook (optional, demonstrates lifecycle rule)
            workbook.Save("PivotItemPositionDemo_out.xlsx");
        }
    }
}
