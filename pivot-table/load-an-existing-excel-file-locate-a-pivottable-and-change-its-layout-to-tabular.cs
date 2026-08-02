// Title: Set PivotTable to Tabular Layout in an Existing Excel Workbook with Aspose.Cells for .NET
// Description: Loads an Excel file, finds the first PivotTable, switches it to tabular form using ShowInTabularForm, refreshes and recalculates the data, then saves the workbook as a new file.
// Keywords: Aspose.Cells C# pivot tabular layout | ShowInTabularForm method | RefreshData Aspose.Cells | CalculateData pivot table | modify pivot table programmatically | load workbook Aspose.Cells | save workbook after pivot change
// Common Searches: Aspose.Cells change pivot layout to tabular | C# set pivot table tabular form | how to refresh pivot after layout change Aspose | find first pivot table in workbook Aspose.Cells | programmatic pivot table layout conversion .NET
// Developer Intent: Apply tabular form to a PivotTable in an existing Excel file and update its data using Aspose.Cells for .NET.
// Use Cases: Standardize financial reports by converting all PivotTables to tabular layout before exporting. | Refresh pivot views after source data modifications to ensure accurate calculations. | Batch process multiple workbooks to enforce a consistent tabular layout across all PivotTables.
// AI Prompts: Generate C# code that iterates through every worksheet in a workbook and sets each PivotTable to tabular form with Aspose.Cells. | Show how to change a PivotTable to tabular layout, refresh, recalculate, and save the file with a timestamped name. | Explain the role of ShowInTabularForm, RefreshData, and CalculateData when updating a PivotTable using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel file, finds the first PivotTable, switches it to tabular form using ShowInTabularForm, refreshes and recalculates the data, then saves the workbook as a new file.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Locate the first PivotTable in the workbook
        PivotTable pivotTable = null;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.PivotTables.Count > 0)
            {
                pivotTable = sheet.PivotTables[0];
                break;
            }
        }

        // If a PivotTable was found, change its layout to Tabular form
        if (pivotTable != null)
        {
            pivotTable.ShowInTabularForm();   // Layout the PivotTable in tabular form
            pivotTable.RefreshData();         // Refresh data from the source
            pivotTable.CalculateData();       // Recalculate the PivotTable
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
