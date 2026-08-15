// Title: Aspose.Cells C# – Enable PivotTable Auto‑Refresh When Workbook Opens
// Description: Shows how to create a workbook, insert sample data, add a PivotTable, and set its RefreshDataOnOpeningFile property to true so the table updates automatically each time the Excel file is opened. The result is saved as PivotTable_RefreshOnOpen.xlsx.
// Keywords: Aspose.Cells | C# PivotTable auto refresh | RefreshDataOnOpeningFile | EnableRefreshOnOpen | pivot table refresh on open | Aspose.Cells .NET example | Excel PivotTable automatic update | Aspose.Cells GitHub demo | pivot table property C# | auto refresh workbook
// Common Searches: Aspose.Cells set PivotTable to refresh on open | C# enable RefreshDataOnOpeningFile for PivotTable | auto refresh Excel PivotTable using Aspose.Cells | how to make PivotTable update when workbook loads .NET | Aspose.Cells PivotTable RefreshDataOnOpeningFile example
// Developer Intent: Configure a PivotTable to refresh automatically when the workbook is opened.
// Use Cases: Sales dashboard that always shows the latest totals without manual refresh. | Inventory report that recalculates stock levels each time the file is opened. | Financial summary sheet that presents up‑to‑date figures on load for stakeholders.
// AI Prompts: Provide C# code that creates an Aspose.Cells PivotTable and enables automatic refresh on workbook open. | Show an example of setting the RefreshDataOnOpeningFile property for a PivotTable and saving the file. | Explain the difference between RefreshDataOnOpeningFile and calling RefreshData manually in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshOnOpenDemo
{
    // Shows how to create a workbook, insert sample data, add a PivotTable, and set its RefreshDataOnOpeningFile property to true so the table updates automatically each time the Excel file is opened. The result is saved as PivotTable_RefreshOnOpen.xlsx.
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
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(2000);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (e.g., add row and data fields)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column

            // Enable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Save the workbook to a file
            workbook.Save("PivotTable_RefreshOnOpen.xlsx");
        }
    }
}
