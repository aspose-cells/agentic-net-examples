// Title: Programmatically hide PivotTable ribbons and field list UI in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Disable the PivotTable field list ribbon for the entire workbook with Aspose.Cells before saving. | Iterate over each worksheet and set EnableWizard and EnableFieldList to false on every PivotTable using C#. | Apply workbook.Settings.HidePivotFieldList = true to suppress all PivotTable UI elements globally in Aspose.Cells.
// Common Searches: how to hide pivot table field list ribbon in Aspose.Cells C# | disable pivot table wizard for all sheets using Aspose.Cells .NET | remove pivot table UI elements before saving Excel file with Aspose.Cells
// Tags: Aspose.Cells suppress pivot UI ribbon | disable pivot wizard C# Aspose | pivot table UI removal Aspose.Cells | save workbook without pivot field list | C# iterate worksheets pivot tables Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, globally hides the PivotTable field list ribbon, iterates through each worksheet to turn off the wizard and field list for every PivotTable, and then saves the workbook with all PivotTable UI elements suppressed.
class DisablePivotRibbonDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Hide the PivotTable field list ribbon for the whole workbook
        workbook.Settings.HidePivotFieldList = true;

        // Iterate through all worksheets and their pivot tables
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (PivotTable pivotTable in sheet.PivotTables)
            {
                // Disable the PivotTable wizard and the field list UI
                pivotTable.EnableWizard = false;
                pivotTable.EnableFieldList = false;
            }
        }

        // Save the workbook with the UI elements disabled
        workbook.Save("output.xlsx");
    }
}
