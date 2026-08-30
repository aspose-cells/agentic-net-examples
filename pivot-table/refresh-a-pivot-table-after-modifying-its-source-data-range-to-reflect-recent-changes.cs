// Title: Refresh an Aspose.Cells PivotTable after modifying its source range using C#
// AI Prompts: Programmatically update source cells and invoke RefreshData and CalculateData on an Aspose.Cells PivotTable in a .NET workbook. | Demonstrate how to recalculate a pivot table after changing underlying data values with Aspose.Cells for C#. | Show the steps to refresh the pivot cache and recompute pivot results after editing source data in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# how to refresh pivot table after editing source cells | Refresh pivot cache programmatically with Aspose.Cells .NET | Recalculate pivot table values after data change using Aspose.Cells API | C# code to update pivot table source range and refresh in Aspose.Cells | Aspose.Cells RefreshData CalculateData example
// Tags: Aspose.Cells refresh pivot cache C# | Aspose.Cells pivot table recalculate data | C# update pivot source range Aspose | Aspose.Cells programmatic pivot refresh .NET | RefreshData CalculateData Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds a pivot table on range A1:B4, modifies source cells B2 and B3, then calls RefreshData and CalculateData to update the pivot table before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table based on the source range A1:B4, placed at E3
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        // Configure the pivot table: Product as row field, Sales as data field
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column index 1 -> Sales

        // Initial calculation to populate the pivot table
        pivot.CalculateData();

        // ----- Modify the source data -----
        sheet.Cells["B2"].PutValue(120); // Change Apple sales from 100 to 120
        sheet.Cells["B3"].PutValue(250); // Change Banana sales from 200 to 250

        // Refresh the pivot table so it reflects the updated source data
        pivot.RefreshData();   // Refreshes the pivot cache from the data source
        pivot.CalculateData(); // Recalculates the pivot table values

        // Save the workbook with the refreshed pivot table
        workbook.Save("RefreshedPivotTable.xlsx");
    }
}
