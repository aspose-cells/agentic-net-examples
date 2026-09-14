// Title: Disable automatic refresh for every PivotTable in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# program with Aspose.Cells that opens an existing .xlsx file, loops through all worksheets and their PivotTables, sets each PivotTable's RefreshDataOnOpeningFile property to false, and saves the modified workbook. | Generate a C# snippet that prevents all PivotTables in a workbook from refreshing when the file is opened, using the Aspose.Cells API to adjust the RefreshDataOnOpeningFile setting.
// Common Searches: Aspose.Cells C# disable pivot table refresh on workbook open | set RefreshDataOnOpeningFile false for all pivot tables in a workbook | how to stop pivot tables from auto‑refreshing when opening an Excel file with Aspose.Cells | improve Excel load speed by turning off pivot table auto refresh using Aspose.Cells .NET | iterate through worksheets to change pivot table settings with Aspose.Cells
// Tags: disable pivot table auto refresh Aspose.Cells | RefreshDataOnOpeningFile property C# | iterate worksheets pivot tables Aspose.Cells | optimize workbook load performance Aspose.Cells | pivot table settings Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The code opens an Excel workbook, iterates over each worksheet and its PivotTables, sets RefreshDataOnOpeningFile to false to stop automatic refresh on opening, and saves the updated file.
class DisablePivotAutoRefresh
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through each PivotTable in the worksheet
            foreach (PivotTable pivotTable in sheet.PivotTables)
            {
                // Disable automatic refresh when the file is opened
                pivotTable.RefreshDataOnOpeningFile = false;
            }
        }

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}
