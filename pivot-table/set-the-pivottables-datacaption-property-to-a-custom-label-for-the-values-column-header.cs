// Title: Changing the values column header of an Aspose.Cells PivotTable with C#
// AI Prompts: Generate a new workbook, add a PivotTable, and replace the default "Values" caption with a custom label by setting the DataFieldHeaderName property using the Aspose.Cells .NET API (C#). | Update an existing PivotTable in a workbook to rename its data field header to a specific string via the PivotTable.DataFieldHeaderName setter in C#.
// Common Searches: Aspose.Cells C# change pivot table values column caption | Set custom header for data field in Excel pivot table using Aspose.Cells | Example of setting pivot table data caption in Aspose.Cells .NET | Rename "Values" column in programmatic pivot table Aspose.Cells C# | How to customize pivot table data caption with Aspose.Cells API
// Tags: Aspose.Cells pivot table custom data caption | C# set pivot table values header | Aspose.Cells modify pivot data field label | Excel pivot table header customization via .NET | Aspose.Cells API rename pivot values column

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, fills it with sample data, adds a PivotTable, assigns a custom caption to the values column using the DataFieldHeaderName property, refreshes and calculates the pivot, and saves the file as PivotTableCustomDataCaption.xlsx.
class SetPivotDataCaption
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Laptop");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(1500);
            sheet.Cells["A4"].PutValue("Phone");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(800);
            sheet.Cells["A5"].PutValue("Phone");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(1200);

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

            // Set a custom caption for the values (data) column header
            pivotTable.DataFieldHeaderName = "Custom Values";

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();   // Correct API to refresh pivot cache
            pivotTable.CalculateData();

            // Save the workbook with the customized pivot table
            workbook.Save("PivotTableCustomDataCaption.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

class Program
{
    static void Main()
    {
        SetPivotDataCaption.Run();
    }
}
