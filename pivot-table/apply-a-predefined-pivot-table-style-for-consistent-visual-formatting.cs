// Title: Apply Built‑In Pivot Table Style (PivotStyleMedium9) Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a pivot table, and assign the built‑in style "PivotStyleMedium9" via the PivotTableStyleName property in Aspose.Cells for .NET, then save the result as an XLSX file.
// Keywords: Aspose.Cells | C# pivot table style | PivotTableStyleName | PivotStyleMedium9 | built‑in pivot style .NET | apply pivot style programmatically | Aspose.Cells example | Excel pivot formatting | consistent pivot appearance | Aspose.Cells API
// Common Searches: Aspose.Cells set pivot table style C# | how to use PivotStyleMedium9 with Aspose.Cells | apply built‑in pivot style .NET | PivotTableStyleName property example | format pivot table using Aspose.Cells
// Developer Intent: Programmatically apply a predefined built‑in style to an Aspose.Cells pivot table for uniform visual formatting.
// Use Cases: Standardize report appearance by applying a corporate‑approved pivot style across all generated workbooks. | Allow end‑users to select a pivot style from a configuration file and apply it at runtime. | Automate the creation of multiple Excel files where each pivot table shares the same branding style.
// AI Prompts: List all built‑in pivot table style names available in Aspose.Cells for .NET. | Generate C# code that changes the PivotTableStyleName of an existing pivot table based on a method parameter. | Show how to apply different built‑in pivot styles conditionally depending on the number of rows in the source data.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotStyleDemo
{
    // Demonstrates how to create a workbook, add a pivot table, and assign the built‑in style "PivotStyleMedium9" via the PivotTableStyleName property in Aspose.Cells for .NET, then save the result as an XLSX file.
    class Program
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

            sheet.Cells["A2"].Value = "Laptop";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Desktop";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Tablet";
            sheet.Cells["B4"].Value = "East";
            sheet.Cells["C4"].Value = 500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Apply a predefined built‑in pivot table style for consistent formatting
            // You can choose any built‑in style name, e.g., "PivotStyleMedium9"
            pivotTable.PivotTableStyleName = "PivotStyleMedium9";

            // Save the workbook with the styled pivot table
            workbook.Save("PivotTableWithPredefinedStyle.xlsx", SaveFormat.Xlsx);
        }
    }
}
