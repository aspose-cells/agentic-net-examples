// Title: Add a Salesperson Report Filter (Page Field) to a Pivot Table with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts sales data, builds a pivot table, adds a Salesperson page field, renames it, refreshes the pivot, opens the filter dialog, and saves the file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# pivot table | report filter page field | salesperson filter | show report filter page | refresh pivot data | rename page field | .NET Excel automation
// Common Searches: Aspose.Cells add page field to pivot table C# | How to show report filter page in Aspose.Cells | Rename pivot table page field Aspose.Cells .NET | Set salesperson as report filter in Excel pivot using Aspose | Refresh and calculate pivot table programmatically Aspose.Cells
// Developer Intent: Insert a Salesperson page filter into a pivot table and programmatically display its filter UI.
// Use Cases: Enable end‑users to filter sales summaries by individual salesperson directly in the generated Excel file. | Create interactive sales dashboards where Region rows are fixed and the Salesperson filter can be changed on demand. | Automate workbook generation that pre‑configures the filter page, refreshes calculations, and opens the filter dialog before distribution.
// AI Prompts: Write C# code with Aspose.Cells that adds a page field named 'Salesperson' to an existing pivot table and opens its filter page. | Show how to rename a pivot table page field, refresh the pivot data, and calculate totals using Aspose.Cells for .NET. | Provide an example that sets a default selected salesperson in the report filter and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsReportFilterDemo
{
    // Creates a workbook, inserts sales data, builds a pivot table, adds a Salesperson page field, renames it, refreshes the pivot, opens the filter dialog, and saves the file using Aspose.Cells for C#.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data with a Salesperson field
            cells["A1"].Value = "Salesperson";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Alice";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Bob";
            cells["B3"].Value = "South";
            cells["C3"].Value = 1500;

            cells["A4"].Value = "Alice";
            cells["B4"].Value = "East";
            cells["C4"].Value = 800;

            cells["A5"].Value = "Charlie";
            cells["B5"].Value = "West";
            cells["C5"].Value = 950;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            // Row field: Region
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            // Data field: Sum of Sales
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            // Page (Report Filter) field: Salesperson
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Salesperson");

            // Optionally rename the page field for clarity
            pivotTable.PageFields[0].Name = "SalespersonFilter";

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Show the report filter page for the Salesperson filter
            pivotTable.ShowReportFilterPageByName("SalespersonFilter");

            // Save the workbook
            workbook.Save("SalesReportWithFilter.xlsx");
        }
    }
}
