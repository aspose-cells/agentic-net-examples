// Title: Enable Excel 2003 compatibility for a PivotTable in Aspose.Cells for .NET to truncate strings over 255 characters before refresh
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds sample data, builds a pivot table, sets IsExcel2003Compatible = true, then refreshes and saves the file. | Show how to enforce the 255‑character string limit in a pivot table by configuring the IsExcel2003Compatible property before calling RefreshData in Aspose.Cells.
// Common Searches: how to enable Excel 2003 compatibility for a pivot table using Aspose.Cells C# | Aspose.Cells truncate pivot table field values longer than 255 characters | set IsExcel2003Compatible property before RefreshData example Aspose.Cells | C# pivot table 255 character limit Aspose.Cells | Aspose.Cells pivot table compatibility mode Excel 2003
// Tags: Aspose.Cells pivot table Excel2003 compatibility | C# set IsExcel2003Compatible property | truncate pivot table strings over 255 characters | refresh pivot data with compatibility mode | create workbook with pivot table Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example demonstrates creating a workbook, populating it with data, adding a pivot table, enabling Excel 2003 compatibility by setting IsExcel2003Compatible to true (which truncates strings longer than 255 characters), refreshing and calculating the pivot data, and saving the workbook as an .xlsx file.
class SetExcel2003CompatibilityDemo
{
    static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate the first worksheet with sample data
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Header row
        dataSheet.Cells["A1"].Value = "Product";
        dataSheet.Cells["B1"].Value = "Description";

        // Data row with a description longer than 255 characters
        string longDescription = new string('x', 300); // 300 characters
        dataSheet.Cells["A2"].Value = "Item1";
        dataSheet.Cells["B2"].Value = longDescription;

        // Add a new worksheet to host the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table (source range A1:B2, destination cell A4)
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B2", "A4", "MyPivotTable");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure pivot fields: Product as row, Description as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Description

        // Enforce Excel 2003 compatibility: strings >255 chars will be truncated
        pivotTable.IsExcel2003Compatible = true;

        // Refresh the pivot cache and calculate the pivot data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotExcel2003Compatible.xlsx");
        Console.WriteLine("Workbook saved successfully.");
    }
}
