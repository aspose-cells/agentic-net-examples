// Title: Disable Excel 2003 compatibility for an Aspose.Cells PivotTable in .NET to retain full text length
// AI Prompts: Generate C# code that creates a workbook, adds a data range with text longer than 255 characters, builds a PivotTable, sets IsExcel2003Compatible = false, refreshes and saves the file. | Show how to prevent 255‑character truncation in an Aspose.Cells PivotTable by disabling Excel 2003 compatibility before calling RefreshData. | Provide a step‑by‑step example of configuring PivotTable.IsExcel2003Compatible in a .NET application and exporting the result to an .xlsx file.
// Common Searches: Aspose.Cells C# set PivotTable.IsExcel2003Compatible false to avoid text cut off | how to keep long description values in Aspose.Cells pivot table | disable Excel 2003 compatibility for pivot cache refresh Aspose.Cells .NET | prevent 255 character limit in Aspose.Cells pivot tables | example of preserving long text in PivotTable using Aspose.Cells for .NET
// Tags: pivot table disable Excel2003 compatibility Aspose.Cells .NET | IsExcel2003Compatible property usage in C# | preserve long text in Aspose.Cells pivot cache | refresh pivot data after disabling compatibility | export pivot table with full text length .xlsx | Aspose.Cells pivot table long description handling

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, adds source data containing a 300‑character description, builds a PivotTable on a separate sheet, disables Excel 2003 compatibility by setting IsExcel2003Compatible to false, refreshes and calculates the pivot cache, and saves the workbook as PivotExcel2003Compatibility.xlsx, ensuring the full text is retained.
class SetPivotExcel2003Compatibility
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare source data worksheet
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";
        dataSheet.Cells["A1"].Value = "Product";
        dataSheet.Cells["B1"].Value = "Description";

        // Add a row with a description longer than 255 characters
        dataSheet.Cells["A2"].Value = "Item1";
        dataSheet.Cells["B2"].Value = new string('X', 300); // 300‑character text

        // Add a worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create a pivot table:
        //   source range: Data!A1:B2
        //   destination cell: A4 (row index 3, column index 0, zero‑based)
        //   table name: "PivotTable1"
        //   useSameSource = false, isXlsClassic = false
        int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B2", 3, 0, "PivotTable1", false, false);
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure the pivot fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Description as data field

        // Disable Excel 2003 compatibility so the long text is not truncated
        pivotTable.IsExcel2003Compatible = false;

        // Refresh the pivot cache and calculate the results
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotExcel2003Compatibility.xlsx");
    }
}
