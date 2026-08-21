// Title: Disable PivotTable auto‑refresh on workbook open with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, builds a PivotTable on A1:B4, sets the EnableRefreshOnOpen (RefreshDataOnOpeningFile) property to false to stop automatic refresh, and saves the file as PivotTable_NoAutoRefresh.xlsx.
// Keywords: Aspose.Cells PivotTable | EnableRefreshOnOpen false | RefreshDataOnOpeningFile | disable pivot auto refresh | C# Excel pivot settings | Aspose.Cells .NET example | prevent pivot refresh on open | Excel workbook performance | static pivot report
// Common Searches: Aspose.Cells set EnableRefreshOnOpen false | how to stop pivot table refresh when opening Excel with Aspose | RefreshDataOnOpeningFile property C# | disable automatic pivot refresh Aspose.Cells | pivot table static data Aspose.Cells .NET
// Developer Intent: Prevent a PivotTable from refreshing its data automatically when the workbook is opened.
// Use Cases: Generate a report workbook where the pivot results must remain unchanged for end users. | Create an Excel template with pre‑calculated pivot values without triggering external data connections. | Improve load time by disabling pivot cache refresh in programmatically generated files.
// AI Prompts: Show C# code that sets EnableRefreshOnOpen (RefreshDataOnOpeningFile) to false for a PivotTable using Aspose.Cells. | Provide a step‑by‑step example of creating a PivotTable and disabling its automatic refresh on open with Aspose.Cells for .NET. | Explain the impact of the EnableRefreshOnOpen property on Excel pivot tables when the file is opened.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds sample data, builds a PivotTable on A1:B4, sets the EnableRefreshOnOpen (RefreshDataOnOpeningFile) property to false to stop automatic refresh, and saves the file as PivotTable_NoAutoRefresh.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(1000);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(2000);
        worksheet.Cells["A4"].PutValue("Orange");
        worksheet.Cells["B4"].PutValue(3000);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // Disable automatic data refresh when the workbook is opened
        pivotTable.RefreshDataOnOpeningFile = false;

        // Save the workbook to a file
        workbook.Save("PivotTable_NoAutoRefresh.xlsx");
    }
}
