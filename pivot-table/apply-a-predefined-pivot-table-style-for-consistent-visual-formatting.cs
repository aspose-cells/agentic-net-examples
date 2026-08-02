// Title: Apply a Built‑In Pivot Table Style with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, fills a small data range, adds a pivot table, assigns row, column and data fields, applies the built‑in style "PivotStyleMedium9" via the PivotTableStyleName property, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# pivot table style | PivotStyleMedium9 | built‑in pivot style .NET | format pivot table programmatically | Excel workbook styling | PivotTableStyleName | sample code | Aspose.Cells tutorial | Excel automation C#
// Common Searches: Aspose.Cells set pivot table style C# | How to use PivotStyleMedium9 in .NET | C# example applying built‑in pivot style | List of built‑in pivot styles Aspose.Cells | Apply consistent formatting to pivot tables programmatically
// Developer Intent: Assign a predefined visual style to a pivot table generated with Aspose.Cells in a C# application.
// Use Cases: Generate a sales‑summary pivot table and apply PivotStyleMedium9 to ensure a uniform look before distribution. | Create multiple regional pivot tables in one workbook and give each the same built‑in style for brand consistency. | Export a styled pivot table to XLSX for downstream reporting or presentation in Microsoft Excel.
// AI Prompts: Show all built‑in pivot table style names available in Aspose.Cells for .NET. | Write C# code that changes the pivot table style based on a user‑selected option using Aspose.Cells. | Explain how to design a custom pivot table style and apply it after the pivot table is created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a new workbook, fills a small data range, adds a pivot table, assigns row, column and data fields, applies the built‑in style "PivotStyleMedium9" via the PivotTableStyleName property, and saves the file as XLSX.
class ApplyPivotTableStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Region";
        sheet.Cells["C1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Laptop";
        sheet.Cells["B2"].Value = "North";
        sheet.Cells["C2"].Value = 1200;

        sheet.Cells["A3"].Value = "Desktop";
        sheet.Cells["B3"].Value = "South";
        sheet.Cells["C3"].Value = 800;

        sheet.Cells["A4"].Value = "Tablet";
        sheet.Cells["B4"].Value = "East";
        sheet.Cells["C4"].Value = 500;

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Apply a predefined built‑in pivot table style for consistent formatting
        pivotTable.PivotTableStyleName = "PivotStyleMedium9";

        // Save the workbook with the styled pivot table
        workbook.Save("PivotTableWithPredefinedStyle.xlsx", SaveFormat.Xlsx);
    }
}
