// Title: Collapse All Column Items in an Aspose.Cells PivotTable (C#) After Refresh
// Description: Demonstrates how to create a workbook, add sample sales data, build a PivotTable with Region rows, Product columns, and Sales values, refresh the table, and programmatically collapse every column‑field item using HideDetail(true) before saving the file.
// Keywords: Aspose.Cells collapse pivot column items | HideDetail true C# | collapse column area after RefreshData | pivot table column collapse .NET | summarized pivot report Aspose.Cells | C# PivotTable hide detail | Aspose.Cells PivotField HideDetail
// Common Searches: how to collapse all column items in Aspose.Cells pivot table C# | Aspose.Cells HideDetail column field after RefreshData | programmatically collapse pivot column area .NET | collapse pivot table columns for summary report | Aspose.Cells pivot table collapse example
// Developer Intent: Programmatically collapse every item in the PivotTable column area after refreshing the data.
// Use Cases: Generate a compact sales summary that shows only region totals while hiding product‑level columns. | Create a dashboard workbook that automatically hides detailed columns after each data refresh. | Provide a report template that presents high‑level totals with column details expandable on demand.
// AI Prompts: Write C# code using Aspose.Cells to add a PivotTable and collapse all column field items after RefreshData. | Show how to enable drill‑down and then hide column details for each PivotField in a PivotTable. | Explain the effect of HideDetail(true) on PivotField objects and how to apply it to multiple column fields.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample sales data, build a PivotTable with Region rows, Product columns, and Sales values, refresh the table, and programmatically collapse every column‑field item using HideDetail(true) before saving the file.
    public class CollapseColumnAreaDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully: CollapsedColumnAreaDemo.xlsx");
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
            // Header row
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            // Data rows
            sheet.Cells["A2"].PutValue("North");
            sheet.Cells["B2"].PutValue("Laptop");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("North");
            sheet.Cells["B3"].PutValue("Phone");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("South");
            sheet.Cells["B4"].PutValue("Laptop");
            sheet.Cells["C4"].PutValue(1500);

            sheet.Cells["A5"].PutValue("South");
            sheet.Cells["B5"].PutValue("Phone");
            sheet.Cells["C5"].PutValue(700);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            // Row area: Region
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            // Column area: Product
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            // Data area: Sum of Sales
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Enable drilldown and show expand/collapse buttons (optional, improves UI)
            pivotTable.EnableDrilldown = true;
            pivotTable.ShowDrill = true;

            // Refresh the pivot table to process the source data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Collapse all items in the column area
            foreach (PivotField columnField in pivotTable.ColumnFields)
            {
                // HideDetail(true) collapses the field for all its items
                columnField.HideDetail(true);
            }

            // Save the workbook with the collapsed column area
            workbook.Save("CollapsedColumnAreaDemo.xlsx");
        }
    }
}
