// Title: Enable AutoSort on a PivotField to sort row items by a related data field in Aspose.Cells for .NET
// AI Prompts: Generate C# code that configures IsAutoSort, IsAscendSort, and AutoSortField on a PivotField so the rows are automatically ordered by the first data field in an Aspose.Cells workbook. | Show how to create a pivot table, add a row field and a data field, and apply automatic ascending sorting to the row field based on the data field using Aspose.Cells for .NET. | Provide a complete example that builds a workbook, populates sample Region and Sales data, adds a pivot table, and enables AutoSort for the row field to order regions by sales values.
// Common Searches: Aspose.Cells C# enable pivot row auto sorting by sales column | How to set AutoSort on a PivotField in Aspose.Cells | Sorting pivot table rows based on a related data field using Aspose.Cells .NET | C# code example for configuring pivot field auto sort in Aspose.Cells | Pivot table row field automatic sorting with Aspose.Cells API
// Tags: pivotfield auto sort aspnet cells | auto sort field index pivot aspnet | c# aspnet cells pivot row sorting | configure pivot table sorting aspnet | aspnet cells pivot auto sort example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotAutoSortDemo
{
    // The sample creates a workbook with Region and Sales data, builds a pivot table, adds Region as a row field and Sales as a data field, enables automatic ascending sorting on the Region field by setting IsAutoSort, IsAscendSort, and AutoSortField, refreshes and calculates the pivot, and saves the result as PivotFieldAutoSortDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Region";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "North";
            cells["B2"].Value = 1200;
            cells["A3"].Value = "South";
            cells["B3"].Value = 1500;
            cells["A4"].Value = "East";
            cells["B4"].Value = 800;
            cells["A5"].Value = "West";
            cells["B5"].Value = 1100;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the row field (Region) and the data field (Sales)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the row field that will be auto‑sorted
            PivotField regionField = pivotTable.RowFields["Region"];

            // Enable auto‑sorting, set ascending order, and sort by the first data field (Sales)
            regionField.IsAutoSort = true;          // Turn on auto sort
            regionField.IsAscendSort = true;        // Ascending order
            regionField.AutoSortField = 0;          // Index of the data field to sort by (Sales)

            // Refresh the pivot table data and calculate results
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldAutoSortDemo.xlsx");
        }
    }
}
