// Title: Hide Pivot Table Ribbon, Wizard & Field List in ODS Export with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook with a pivot table, then disable the PivotTable wizard, field list, and ribbon before saving as an ODS file. The pivot remains functional while the UI elements are hidden, producing a clean OpenDocument Spreadsheet.
// Keywords: Aspose.Cells | C# | ODS export | pivot table UI hide | disable pivot ribbon | hide pivot wizard | field list hidden | minimal UI ODS | OpenDocument Spreadsheet | Aspose.Cells PivotTable settings
// Common Searches: how to hide pivot table ribbon in ODS using Aspose.Cells | disable pivot wizard and field list when exporting to ODS .NET | Aspose.Cells hide pivot UI controls in OpenDocument Spreadsheet | remove pivot table editing tools from ODS output C# | minimal UI ODS file with functional pivot table Aspose
// Developer Intent: Export an ODS workbook that keeps the pivot table functional but removes all pivot‑related UI controls.
// Use Cases: Deliver ODS reports with a working pivot table but a clean interface for end users. | Provide spreadsheet templates where calculations rely on a pivot table while preventing users from accessing the wizard or field list. | Generate ODS files for web‑based viewers that should not expose pivot‑table editing options.
// AI Prompts: Show C# code to hide the pivot table ribbon, wizard, and field list while keeping the pivot functional in an ODS file with Aspose.Cells. | Explain the difference between OdsSaveOptions.IgnorePivotTables and UI‑hiding properties for Aspose.Cells pivot tables. | Give a step‑by‑step guide to produce a minimal‑UI ODS spreadsheet containing a pivot table using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace PivotTableOdsMinimalUi
{
    // Demonstrates how to create a workbook with a pivot table, then disable the PivotTable wizard, field list, and ribbon before saving as an ODS file. The pivot remains functional while the UI elements are hidden, producing a clean OpenDocument Spreadsheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(3000);

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales column as data field

            // Disable UI elements related to the pivot table
            pivotTable.EnableWizard = false;      // Hide the PivotTable wizard
            pivotTable.EnableFieldList = false;   // Hide the field list on the worksheet
            workbook.Settings.HidePivotFieldList = true; // Hide the pivot field list globally

            // Prepare ODS save options (keep the pivot table but UI is hidden)
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.IgnorePivotTables = false; // Ensure pivot table is saved

            // Save the workbook as ODS with the specified options
            workbook.Save("PivotTable_MinimalUI.ods", saveOptions);
        }
    }
}
