// Title: How to hide pivot table field list and wizard ribbons when saving a workbook as ODS using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a pivot table, disables the field‑list and wizard ribbons, and saves the workbook as an ODS file with Aspose.Cells. | Write a snippet that sets workbook.Settings.HidePivotFieldList and pivot.EnableWizard = false to suppress all pivot‑related ribbons before ODS export. | Provide an example showing how to turn off the pivot table UI ribbons for a specific pivot table and export the result to ODS in a .NET application.
// Common Searches: asp.net aspose.cells hide pivot field list ribbon when exporting to ods | c# disable pivot table wizard ribbon in ods file | how to suppress pivot UI ribbons in ODS using Aspose.Cells | remove pivot table ribbon from generated ODS document c#
// Tags: hide pivot field list Aspose.Cells | disable pivot wizard ODS export | pivot table ribbon suppression .NET | Aspose.Cells ODS pivot settings | pivot UI ribbon removal C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Ods;

// The example creates a workbook, adds sample data and a pivot table, then disables the field‑list and wizard ribbons via workbook and pivot properties before saving the file as an ODS document.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Food");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["A3"].PutValue("Clothing");
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["A4"].PutValue("Electronics");
        sheet.Cells["B4"].PutValue(1500);

        // Add a pivot table to the worksheet
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount

        // Disable UI elements (ribbons) related to the pivot table
        // Hide the field list (ribbon) in the workbook
        workbook.Settings.HidePivotFieldList = true;
        // Disable the PivotTable wizard and field list for this pivot table
        pivot.EnableWizard = false;
        pivot.EnableFieldList = false;

        // Save the workbook as ODS using default options (no need to modify OdsSaveOptions)
        workbook.Save("PivotTable_NoRibbons.ods", SaveFormat.Ods);
    }
}
