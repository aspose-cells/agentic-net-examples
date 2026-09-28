// Title: Enable RefreshDataOnOpeningFile for an Aspose.Cells PivotTable so it updates automatically when the workbook opens (C#)
// AI Prompts: Set the PivotTable.RefreshDataOnOpeningFile property to true in a C# Aspose.Cells workbook. | Create a new workbook, add a PivotTable, and configure it to refresh automatically on file open using Aspose.Cells for .NET. | Modify existing Aspose.Cells C# code to turn on automatic PivotTable data refresh when the workbook is opened.
// Common Searches: Aspose.Cells C# how to make a pivot table refresh on workbook open | set RefreshDataOnOpeningFile property for PivotTable in Aspose.Cells .NET | automatic pivot table refresh when opening Excel file using Aspose.Cells | C# example enabling pivot table auto refresh with Aspose.Cells | Aspose.Cells enable pivot table data refresh on file load
// Tags: PivotTable.RefreshDataOnOpeningFile property Aspose.Cells | auto refresh pivot table C# Aspose.Cells | enable pivot table refresh on workbook open .NET | Aspose.Cells pivot table automatic refresh Excel | set pivot table refresh flag Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, adds sample data, inserts a PivotTable, sets PivotTable.RefreshDataOnOpeningFile to true so the table refreshes automatically when the file is opened, and saves the workbook as PivotRefreshOnOpen.xlsx.
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
        sheet.Cells["B4"].PutValue(1500);

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // Enable automatic refresh when the workbook is opened
        pivotTable.RefreshDataOnOpeningFile = true;

        // Save the workbook to a file
        workbook.Save("PivotRefreshOnOpen.xlsx");
    }
}
