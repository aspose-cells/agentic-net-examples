// Title: Aspose.Cells for .NET – Hide PivotTable Report Filter Pane (ShowReportFilter = false)
// Description: Creates a workbook, adds sample data, builds a PivotTable on A1:B5, assigns row and data fields, sets PivotTable.ShowReportFilter = false to hide the report‑filter area, then saves the file.
// Keywords: Aspose.Cells hide pivot report filter | ShowReportFilter false C# | PivotTable hide filter pane .NET | Aspose.Cells PivotTable display options | remove report filter area Excel | Aspose.Cells PivotTable properties
// Common Searches: how to hide pivot report filter Aspose.Cells | ShowReportFilter property example C# | disable report filter pane in generated Excel | Aspose.Cells hide PivotTable filter area | C# hide PivotTable report filter
// Developer Intent: Hide the report‑filter area of a PivotTable generated with Aspose.Cells for .NET.
// Use Cases: Generate cleaner Excel reports without the filter pane. | Create printable PivotTables that fit page layouts. | Build dashboards where the filter area is unnecessary.
// AI Prompts: Write C# code using Aspose.Cells to create a PivotTable and set ShowReportFilter = false before saving. | Show an example that adds a row field and a data field to a PivotTable and hides the report filter area. | Demonstrate how to toggle the ShowReportFilter property after refreshing PivotTable data.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // Creates a workbook, adds sample data, builds a PivotTable on A1:B5, assigns row and data fields, sets PivotTable.ShowReportFilter = false to hide the report‑filter area, then saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
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

                // Add a pivot table based on the data range A1:B5, placed at D3
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure fields: Category as row field, Amount as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh the pivot cache and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_HideReportFilter.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
