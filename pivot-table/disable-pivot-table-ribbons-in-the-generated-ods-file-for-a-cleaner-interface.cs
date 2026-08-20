// Title: Hide Pivot Table Ribbons and Field List in ODS using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, build a pivot table, then hide the PivotTable field‑list ribbon (Workbook.Settings.HidePivotFieldList) and disable the wizard and field list for a specific pivot (PivotTable.EnableWizard, PivotTable.EnableFieldList) before saving the file as ODS with OdsSaveOptions.
// Keywords: Aspose.Cells hide pivot ribbon | disable pivot wizard C# | remove pivot field list ODS | Workbook.Settings.HidePivotFieldList | PivotTable.EnableWizard | PivotTable.EnableFieldList | Aspose.Cells ODS export | C# pivot table UI settings | .NET generate clean ODS
// Common Searches: how to hide pivot table ribbon in ODS with Aspose.Cells | disable pivot table wizard and field list C# | Aspose.Cells hide pivot UI before saving | remove pivot ribbons from exported ODS file | C# code to hide pivot field list in workbook
// Developer Intent: Programmatically suppress all PivotTable UI elements—ribbons, wizard, and field list—so the exported ODS file presents a clean, read‑only interface.
// Use Cases: Creating ODS reports that show pivot results without exposing editing tools to end users. | Generating read‑only pivot tables for shared or automated processing pipelines. | Delivering streamlined ODS files for web or mobile viewers where UI clutter must be minimized.
// AI Prompts: Generate C# code with Aspose.Cells that hides the pivot field‑list ribbon for the whole workbook and disables the wizard and field list for a specific pivot table before saving as ODS. | Show how to programmatically turn off PivotTable UI elements (ribbons, wizard, field list) in an ODS file using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add sample data, build a pivot table, then hide the PivotTable field‑list ribbon (Workbook.Settings.HidePivotFieldList) and disable the wizard and field list for a specific pivot (PivotTable.EnableWizard, PivotTable.EnableFieldList) before saving the file as ODS with OdsSaveOptions.
public class DisablePivotTableRibbons
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("Food");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["A3"].PutValue("Drink");
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["A4"].PutValue("Food");
        worksheet.Cells["B4"].PutValue(150);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

        // ----- Disable pivot table ribbons / UI elements -----
        // Hide the PivotTable field list ribbon for the whole workbook
        workbook.Settings.HidePivotFieldList = true;

        // Disable the PivotTable wizard and field list for this specific pivot table
        pivotTable.EnableWizard = false;
        pivotTable.EnableFieldList = false;
        // ----------------------------------------------------

        // Save the workbook as ODS using OdsSaveOptions
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        workbook.Save("PivotTable_NoRibbons.ods", saveOptions);
    }
}
