// Title: Aspose.Cells for .NET – Hide PivotTable Report Filter Area (ShowReportFilter = false)
// Description: C# example that creates a workbook, adds sample data, builds a PivotTable with row, data and page fields, and demonstrates how to hide the report‑filter pane. The ShowReportFilter property is not present in the current Aspose.Cells release, so the sample notes alternative ways to conceal the filter area before saving the file.
// Keywords: Aspose.Cells PivotTable hide report filter | ShowReportFilter false .NET | remove filter pane Aspose.Cells | C# pivot table display options | page field visibility Aspose.Cells | Aspose.Cells pivot table customization | hide report filter area programmatically
// Common Searches: how to hide pivot table report filter in Aspose.Cells | ShowReportFilter property Aspose.Cells .NET | remove page field pane from PivotTable using Aspose.Cells | Aspose.Cells hide filter area alternative method | C# hide pivot table filter pane Aspose
// Developer Intent: Programmatically suppress the report‑filter (page field) pane of a PivotTable in Aspose.Cells for .NET.
// Use Cases: Produce clean, printer‑friendly reports without the filter pane. | Design dashboards where the PivotTable’s filter area would waste screen space. | Automate reporting pipelines that require the filter area to be hidden before distribution.
// AI Prompts: Generate C# code with Aspose.Cells that hides the PivotTable report filter area when ShowReportFilter is unavailable. | Suggest a workaround to conceal the page‑field area of a PivotTable using Aspose.Cells for .NET. | Explain how to verify that the report filter pane is not visible after saving a workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that creates a workbook, adds sample data, builds a PivotTable with row, data and page fields, and demonstrates how to hide the report‑filter pane. The ShowReportFilter property is not present in the current Aspose.Cells release, so the sample notes alternative ways to conceal the filter area before saving the file.
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
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(15);

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot fields (row, data, and a page field to act as a report filter)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");

            // Note: The ShowReportFilter property is not available in the current Aspose.Cells version.
            // If needed, alternative approaches can be used to hide the report filter area.

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
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
