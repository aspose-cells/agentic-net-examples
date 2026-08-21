// Title: C# – Apply PivotTableStyleMedium9 to a PivotTable with Aspose.Cells
// Description: The sample builds a workbook, adds sample sales rows, creates a pivot table, assigns row, column and data fields, calculates the pivot data, and then applies the built‑in PivotTableStyleMedium9 format through the PivotTableStyleType property before saving the file.
// Keywords: Aspose.Cells C# pivot table style | PivotTableStyleMedium9 example | set pivot table format programmatically | Excel pivot styling Aspose | apply built‑in pivot style .NET | C# Aspose.Cells workbook formatting | quick pivot table styling
// Common Searches: how to set PivotTableStyleMedium9 in Aspose.Cells | C# code for applying built‑in pivot style | Aspose.Cells pivot table formatting tutorial | assign default style to Excel pivot table using .NET | programmatic pivot table styling Aspose
// Developer Intent: Programmatically assign the PivotTableStyleMedium9 format to a pivot table.
// Use Cases: Generate a sales analysis workbook and give its pivot table a polished look instantly. | Enforce a consistent visual theme across all pivot tables in automated reporting pipelines. | Speed up workbook creation by applying a ready‑made style instead of manual formatting.
// AI Prompts: Provide C# code that sets PivotTableStyleMedium9 on a pivot table after data calculation with Aspose.Cells. | Show how to choose a different built‑in pivot style based on a runtime condition in Aspose.Cells for .NET. | List all built‑in pivot table styles available in Aspose.Cells and demonstrate selecting one programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample builds a workbook, adds sample sales rows, creates a pivot table, assigns row, column and data fields, calculates the pivot data, and then applies the built‑in PivotTableStyleMedium9 format through the PivotTableStyleType property before saving the file.
class AssignPivotTableStyle
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Region";
        sheet.Cells["C1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Bike";
        sheet.Cells["B2"].Value = "North";
        sheet.Cells["C2"].Value = 5000;

        sheet.Cells["A3"].Value = "Car";
        sheet.Cells["B3"].Value = "South";
        sheet.Cells["C3"].Value = 12000;

        sheet.Cells["A4"].Value = "Truck";
        sheet.Cells["B4"].Value = "East";
        sheet.Cells["C4"].Value = 8000;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Calculate the pivot data
        pivotTable.CalculateData();

        // Assign the built‑in style PivotTableStyleMedium9
        pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

        // Save the workbook
        workbook.Save("PivotTableWithMedium9Style.xlsx");
    }
}
