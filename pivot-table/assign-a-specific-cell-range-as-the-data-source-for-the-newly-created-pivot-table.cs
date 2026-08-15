// Title: Change Pivot Table Data Source Range Using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add a pivot table, then reassign its source to a different cell range (e.g., B2:C5) with PivotTable.ChangeDataSource, refresh the data, recalculate, and save the file.
// Keywords: Aspose.Cells pivot table change data source | C# PivotTable.ChangeDataSource example | update pivot source range Aspose.Cells | .NET Excel pivot table data source | dynamic pivot table range C#
// Common Searches: Aspose.Cells how to change pivot table source range C# | PivotTable.ChangeDataSource usage .NET | replace pivot table data source Aspose.Cells | C# set new range for existing pivot table
// Developer Intent: Replace the source range of an existing pivot table with another cell range programmatically.
// Use Cases: Create a pivot from a full dataset, then narrow the analysis to a subset without rebuilding the layout. | Allow end‑users to select a different data block (e.g., filtered rows) and update the pivot on the fly. | Adjust the pivot source when the underlying table expands or contracts, then refresh automatically.
// AI Prompts: Generate C# code that creates a pivot table with Aspose.Cells and then changes its data source to B2:C5. | Explain step‑by‑step how to use PivotTable.ChangeDataSource to point a pivot table to a new range and recalculate it. | Show an example of dynamically updating a pivot table's source range based on user input using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to create a workbook, add a pivot table, then reassign its source to a different cell range (e.g., B2:C5) with PivotTable.ChangeDataSource, refresh the data, recalculate, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (A1:C5)
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Region");
        worksheet.Cells["C1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue("North");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue("South");
        worksheet.Cells["C3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue("East");
        worksheet.Cells["C4"].PutValue(300);
        worksheet.Cells["A5"].PutValue("D");
        worksheet.Cells["B5"].PutValue("West");
        worksheet.Cells["C5"].PutValue(400);

        // Add a pivot table with an initial source (will be changed later)
        int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "MyPivotTable");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Assign a specific cell range as the new data source using ChangeDataSource
        // Example: use the range B2:C5 as the data source
        string[] newDataSource = new string[] { "B2:C5" };
        pivotTable.ChangeDataSource(newDataSource);

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Refresh and calculate the pivot table to apply changes
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTableAssignDataSource.xlsx");
    }
}
