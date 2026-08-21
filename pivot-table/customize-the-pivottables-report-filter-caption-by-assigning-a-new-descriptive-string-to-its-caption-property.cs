// Title: Set a custom caption for a PivotTable report filter (page field) in C# with Aspose.Cells
// Description: Shows how to create a workbook, add sample sales data, build a PivotTable, place the "Region" field as a page (report filter) field, and assign a user‑friendly caption (e.g., "Select Region") by updating the PageField's Name property, then refresh and save the workbook.
// Keywords: Aspose.Cells | C# PivotTable | report filter caption | page field name | custom filter label | set pivot filter caption | Aspose.Cells PivotTable API | .NET spreadsheet library | programmatic Excel pivot | Excel automation
// Common Searches: Aspose.Cells change PivotTable filter caption C# | set custom name for page field Aspose.Cells | how to rename PivotTable report filter using .NET | Aspose.Cells PivotTable page field label example | C# code to customize PivotTable filter dropdown text
// Developer Intent: Programmatically assign a descriptive caption to a PivotTable report filter (page field).
// Use Cases: Provide a clear label such as "Select Region" for end‑users when generating sales dashboards. | Standardize filter captions across multiple automatically created PivotTables. | Improve readability of Excel reports produced by a .NET reporting service.
// AI Prompts: Generate C# code that sets a custom caption for a PivotTable page field using Aspose.Cells. | Explain how to rename multiple report filter fields in a PivotTable with Aspose.Cells for .NET. | Show the steps to refresh and recalculate a PivotTable after changing its filter caption.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add sample sales data, build a PivotTable, place the "Region" field as a page (report filter) field, and assign a user‑friendly caption (e.g., "Select Region") by updating the PageField's Name property, then refresh and save the workbook.
    public class PivotTableReportFilterCaptionDemo
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
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 600;

            sheet.Cells["A5"].Value = "Vegetable";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 900;

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table areas
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region"); // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // Data field

            // Add a report filter (page field) and customize its caption
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");
            // The 'Name' of the page field acts as the caption displayed for the filter
            pivotTable.PageFields[0].Name = "Select Region";

            // Refresh the pivot cache and calculate the pivot table data
            pivotTable.RefreshData();   // Correct API to refresh data source
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableReportFilterCaptionDemo.xlsx");
        }
    }
}
