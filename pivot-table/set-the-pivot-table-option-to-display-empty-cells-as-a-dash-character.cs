// Title: Aspose.Cells C# Pivot Table – Show a Dash for Empty Cells
// Description: Creates a workbook, adds sample data with blanks, builds a pivot table, enables DisplayNullString, sets NullString to "-", refreshes the table and saves the file so empty cells appear as a dash.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | DisplayNullString | NullString | empty cell placeholder | dash character | custom null string | Excel export
// Common Searches: Aspose.Cells pivot table display dash for null values | C# set NullString in Aspose.Cells pivot table | DisplayNullString property example Aspose.Cells | replace empty pivot cells with '-' in .NET | how to show dash instead of blank in Excel pivot using Aspose
// Developer Intent: Configure a pivot table to replace null or blank values with a dash character.
// Use Cases: Generate reports where missing categories are shown as "-" for better readability. | Create Excel workbooks with pivot tables that use a custom placeholder for empty data cells. | Automate export of pivot tables where consistent formatting of empty values is required.
// AI Prompts: Write C# code with Aspose.Cells to enable DisplayNullString and set NullString to "-" for a pivot table. | Explain the effect of DisplayNullString and NullString on empty cell rendering in an Aspose.Cells pivot table. | Provide step‑by‑step instructions to build a pivot table in C# and customize empty cells to display a dash.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook, adds sample data with blanks, builds a pivot table, enables DisplayNullString, sets NullString to "-", refreshes the table and saves the file so empty cells appear as a dash.
    class DisplayDashForEmptyCells
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (data source)
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data with some empty cells
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("");   // Empty category cell
            dataSheet.Cells["B4"].PutValue(30);
            dataSheet.Cells["A5"].PutValue("C");
            dataSheet.Cells["B5"].PutValue(null); // Empty value cell

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table based on the data range
            int pivotIndex = pivotSheet.PivotTables.Add("=Sheet1!A1:B5", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value

            // Set the pivot table to display a custom string for null/empty cells
            pivotTable.DisplayNullString = true;   // Enable custom null string
            pivotTable.NullString = "-";           // Use dash character for empty cells

            // Refresh data and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_DisplayDashForEmptyCells.xlsx");
        }
    }
}
