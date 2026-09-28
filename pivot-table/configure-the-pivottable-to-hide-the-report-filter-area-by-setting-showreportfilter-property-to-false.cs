// Title: Hide the report filter area of a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a pivot table, and disables the report filter UI by setting ShowReportFilter = false. | Update an existing Aspose.Cells pivot‑table example to turn off the report filter pane without altering row or data fields. | Show how to programmatically hide the report filter section of a PivotTable after refreshing data with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# remove report filter section from generated pivot table | Set ShowReportFilter property to false in Aspose.Cells pivot table example | How to remove report filter area from a PivotTable using Aspose.Cells for .NET | C# Aspose.Cells pivot table UI options hide filter pane | Disable report filter area in Excel pivot table created with Aspose.Cells
// Tags: Aspose.Cells pivot table ShowReportFilter | hide report filter UI Aspose.Cells | C# pivot table display options Aspose.Cells | Excel pivot table filter pane suppression Aspose.Cells | Aspose.Cells workbook pivot table configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // The example creates a new workbook, fills it with sample data, adds a pivot table on range A1:B5, assigns 'Category' as a row field and 'Amount' as a data field, refreshes and calculates the pivot, sets ShowReportFilter to false to hide the filter pane, and saves the file as PivotTable_HideReportFilter.xlsx.
    class HideReportFilterArea
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Drink");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Food");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Drink");
                sheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh the pivot cache and calculate the pivot table
                pivotTable.RefreshData();   // Correct API call
                pivotTable.CalculateData();

                // Save the workbook to a file
                string outputPath = "PivotTable_HideReportFilter.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
