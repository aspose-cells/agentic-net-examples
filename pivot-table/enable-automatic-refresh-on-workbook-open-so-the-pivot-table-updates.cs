// Title: Aspose.Cells C# – Enable PivotTable auto‑refresh when workbook opens
// Description: The sample builds a workbook, inserts sample data, creates a PivotTable, maps Product to rows and Sales to values, activates the RefreshDataOnOpeningFile flag, and saves the file as PivotTable_AutoRefreshOnOpen.xlsx so the pivot updates each time the workbook is opened.
// Keywords: Aspose.Cells | C# | PivotTable | RefreshDataOnOpeningFile | auto refresh | workbook open | Excel pivot update | Aspose.Cells example | pivot table settings
// Common Searches: Aspose.Cells set pivot to refresh on open | C# pivot table auto refresh Aspose | RefreshDataOnOpeningFile usage | How to make Excel pivot update automatically with Aspose.Cells | Enable workbook open refresh for pivot tables .NET
// Developer Intent: Configure a PivotTable so it refreshes its source data automatically whenever the workbook is opened.
// Use Cases: Distribute a sales dashboard that always shows the latest figures without manual refresh. | Create a template for monthly reporting where the pivot recalculates on each open. | Maintain a shared analytics workbook in a collaborative environment ensuring consistent, up‑to‑date pivot results.
// AI Prompts: Generate C# code using Aspose.Cells to add a PivotTable and turn on RefreshDataOnOpeningFile. | Show how to apply RefreshDataOnOpeningFile to all PivotTables in an existing workbook with Aspose.Cells. | Explain how the RefreshDataOnOpeningFile property affects file size and opening performance in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshOnOpen
{
    // The sample builds a workbook, inserts sample data, creates a PivotTable, maps Product to rows and Sales to values, activates the RefreshDataOnOpeningFile flag, and saves the file as PivotTable_AutoRefreshOnOpen.xlsx so the pivot updates each time the workbook is opened.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(1200);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(850);
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(950);

            // Add a pivot table based on the data range
            int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

            // Configure the pivot table (Product as row, Sales as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Column index 1 -> Sales

            // Enable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Save the workbook (the setting will be persisted)
            workbook.Save("PivotTable_AutoRefreshOnOpen.xlsx");
        }
    }
}
