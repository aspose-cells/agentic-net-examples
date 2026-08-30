// Title: Add a salesperson report filter (page field) to a pivot table using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a pivot table from a data range and adds a page field named Salesperson as a report filter with Aspose.Cells. | Demonstrate how to rename a pivot table page field and programmatically open its filter dialog using Aspose.Cells for .NET. | Outline the steps to refresh, calculate, and save a workbook containing a pivot table with a salesperson filter to an .xlsx file.
// Common Searches: how to add a page field as a report filter in an Aspose.Cells pivot table C# | Aspose.Cells rename pivot table filter field and display filter page example | C# code to create a salesperson filter for a pivot table using Aspose.Cells | saving a workbook with a pivot table and report filter using Aspose.Cells .NET | display report filter dialog for a pivot table programmatically Aspose.Cells
// Tags: pivot table page field Aspose.Cells C# | rename pivot filter field Aspose.Cells | show report filter dialog Aspose.Cells | save workbook with pivot filter Aspose.Cells | salesperson report filter pivot table Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, fills it with sample sales data, adds a pivot table, sets Region as a row field and Sales as a data field, adds Salesperson as a page (report filter) field, renames the filter to "SalespersonFilter", opens the filter page, refreshes and calculates the pivot data, and saves the workbook as SalesPivotReportFilter.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample sales data
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

        // Add Region as a row field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");

        // Add Sales as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add Salesperson as a page (report filter) field
        pivotTable.AddFieldToArea(PivotFieldType.Page, "Salesperson");
        // Optionally rename the filter field for clarity
        pivotTable.PageFields[0].Name = "SalespersonFilter";

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Show the report filter page for the salesperson filter
        pivotTable.ShowReportFilterPageByName("SalespersonFilter");

        // Save the workbook with the pivot table and filter page
        workbook.Save("SalesPivotReportFilter.xlsx");
    }
}
