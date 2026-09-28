// Title: How to hide a specific row PivotField in an Aspose.Cells for .NET pivot table by hiding its items
// AI Prompts: Generate C# code using Aspose.Cells that retrieves a row field from a pivot table and calls HideItem on each of its items to hide the field. | Show the steps to programmatically hide a PivotField in an Excel workbook with Aspose.Cells, then refresh and recalculate the pivot table.
// Common Searches: Aspose.Cells C# hide row field in pivot table example | programmatically hide a PivotField in Excel using Aspose.Cells .NET | how to set PivotField items hidden in Aspose.Cells pivot table | C# code to hide a specific field in an Aspose.Cells pivot table and refresh data | Aspose.Cells hide product field in pivot table rows
// Tags: Aspose.Cells hide PivotField items | C# hide pivot table field programmatically | Aspose.Cells refresh pivot after hiding field | Excel pivot table hide field using Aspose | Aspose.Cells set PivotField hidden property

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHideFieldDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, places the 'Product' field in the row area, iterates through its items calling HideItem(true) to hide the field, refreshes and recalculates the pivot, and saves the file as PivotFieldHiddenDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["A5"].PutValue("Orange");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(1300);
            sheet.Cells["B5"].PutValue(1700);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Product" field to the row area and "Sales" to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the row field we just added
            PivotField productField = pivotTable.RowFields[0];

            // Hide all items of the "Product" field.
            // This effectively hides the entire field from the pivot view.
            for (int i = 0; i < productField.ItemCount; i++)
            {
                productField.HideItem(i, true);
            }

            // Refresh and calculate the pivot table after modifications
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the hidden pivot field
            workbook.Save("PivotFieldHiddenDemo.xlsx");
        }
    }
}
