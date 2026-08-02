// Title: Hide PivotTable UI (field list ribbon, wizard) in Aspose.Cells for .NET
// Description: Demonstrates how to programmatically suppress the PivotTable field‑list ribbon, wizard, and field list in a workbook created with Aspose.Cells for .NET before saving it as an Excel file.
// Keywords: Aspose.Cells hide pivot field list | disable pivot ribbon .NET | turn off pivot wizard Aspose.Cells | remove pivot UI Excel | HidePivotFieldList | PivotTable.EnableWizard | PivotTable.EnableFieldList | C# Aspose.Cells pivot table
// Common Searches: how to hide pivot field list ribbon using Aspose.Cells C# | disable pivot table wizard in generated Excel with Aspose.Cells | remove pivot UI elements before saving workbook | Aspose.Cells hide pivot UI Excel output | C# programmatically turn off pivot table ribbons
// Developer Intent: Programmatically hide the PivotTable field‑list ribbon, wizard, and field list in an Aspose.Cells workbook to deliver a clean Excel UI.
// Use Cases: Create a reporting workbook where the PivotTable opens without the field‑list ribbon, preventing accidental layout changes. | Distribute Excel templates that show data only, with the PivotTable wizard and field list disabled for end users. | Automate Excel exports containing PivotTables while ensuring a minimal UI for non‑technical recipients.
// AI Prompts: Show C# code that disables the PivotTable field‑list ribbon, wizard, and field list using Aspose.Cells before saving the workbook. | Provide an Aspose.Cells example to turn off PivotTable UI elements for a specific pivot table in .NET. | Explain how to suppress all PivotTable UI components in an Aspose.Cells workbook to reduce clutter.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRibbonDisable
{
    // Demonstrates how to programmatically suppress the PivotTable field‑list ribbon, wizard, and field list in a workbook created with Aspose.Cells for .NET before saving it as an Excel file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Drink");
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(1500);
            sheet.Cells["A5"].PutValue("Drink");
            sheet.Cells["B5"].PutValue(900);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field

            // Disable pivot table UI ribbons/field list to reduce UI clutter
            // This hides the PivotTable field list ribbon when the workbook is opened in Excel
            workbook.Settings.HidePivotFieldList = true;

            // Optionally also disable the PivotTable wizard and field list for the specific pivot table
            pivotTable.EnableWizard = false;
            pivotTable.EnableFieldList = false;

            // Save the workbook
            workbook.Save("PivotTable_RibbonDisabled.xlsx");
        }
    }
}
