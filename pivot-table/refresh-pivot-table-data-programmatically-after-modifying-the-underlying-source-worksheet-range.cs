// Title: Programmatically refresh an Aspose.Cells PivotTable after updating source worksheet cells in C#
// AI Prompts: Generate C# code that changes cell values in a worksheet and then calls Worksheet.RefreshPivotTables to update the PivotTable using Aspose.Cells. | Show how to recalculate a PivotTable after modifying its source data by invoking PivotTable.CalculateData followed by Worksheet.RefreshPivotTables in a .NET application. | Provide an example that creates a workbook, adds a PivotTable, edits the source range, and refreshes all pivot tables to reflect the new values.
// Common Searches: Aspose.Cells C# refresh pivot table after changing cell values | how to update pivot table data programmatically with Aspose.Cells .NET | Worksheet.RefreshPivotTables method example for recalculating pivot tables | refresh all pivot tables in a workbook using Aspose.Cells C# | recalculate pivot table after source data edit Aspose.Cells
// Tags: Aspose.Cells refresh pivot tables C# | Worksheet.RefreshPivotTables usage | update source data recalculate pivot Aspose.Cells | programmatic pivot table refresh .NET | PivotTable.CalculateData after cell edit

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, fills source data, adds a PivotTable, modifies the source cells, calls Worksheet.RefreshPivotTables (and optionally PivotTable.CalculateData) to recalculate the PivotTable, and saves the result as RefreshedPivot.xlsx.
public class RefreshPivotTableDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["A4"].PutValue("Orange");
            worksheet.Cells["B4"].PutValue(300);

            // Add a pivot table based on the source range A1:B4, place it at D1
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate the initial pivot data
            pivotTable.CalculateData();

            // Modify the underlying source data
            worksheet.Cells["B2"].PutValue(150); // Update Apple sales
            worksheet.Cells["B3"].PutValue(250); // Update Banana sales

            // Refresh all pivot tables in the worksheet to reflect the changes
            worksheet.RefreshPivotTables();

            // Save the workbook with the refreshed pivot table
            workbook.Save("RefreshedPivot.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        RefreshPivotTableDemo.Run();
    }
}
