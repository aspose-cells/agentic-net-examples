// Title: Disable PivotTable Auto‑Refresh on Workbook Open with Aspose.Cells for .NET
// Description: Loads an existing Excel file, iterates through every worksheet and each PivotTable, sets RefreshDataOnOpeningFile to false to stop automatic refresh when the workbook is opened, and saves the updated file. This reduces load time for large workbooks and keeps pivot data static until manually refreshed.
// Keywords: Aspose.Cells disable pivot auto refresh | RefreshDataOnOpeningFile false | pivot table performance .NET | prevent pivot refresh on open | Excel workbook load optimization | iterate worksheets Aspose | C# Aspose.Cells PivotTable settings
// Common Searches: how to turn off pivot table auto refresh using Aspose.Cells | Aspose.Cells .NET disable RefreshDataOnOpeningFile | speed up Excel opening by disabling pivot refresh | set pivot tables not to refresh on file open C# | batch disable pivot auto refresh Aspose
// Developer Intent: Programmatically prevent every PivotTable in an Excel workbook from refreshing automatically when the file is opened.
// Use Cases: Distribute a report where pivot data must stay unchanged until the user decides to refresh. | Improve startup performance of large workbooks that contain many PivotTables. | Automate processing of multiple Excel files to ensure consistent pivot behavior across all documents.
// AI Prompts: Generate C# code that disables automatic refresh for all PivotTables in a workbook using Aspose.Cells. | Create a reusable method that takes an input and output path, turns off PivotTable auto refresh, and saves the file. | Explain how to verify that each PivotTable's RefreshDataOnOpeningFile property is set to false after modification.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an existing Excel file, iterates through every worksheet and each PivotTable, sets RefreshDataOnOpeningFile to false to stop automatic refresh when the workbook is opened, and saves the updated file. This reduces load time for large workbooks and keeps pivot data static until manually refreshed.
class DisablePivotAutoRefresh
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Loop through each PivotTable in the worksheet
            foreach (PivotTable pivotTable in worksheet.PivotTables)
            {
                // Disable automatic refresh when the file is opened
                pivotTable.RefreshDataOnOpeningFile = false;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
