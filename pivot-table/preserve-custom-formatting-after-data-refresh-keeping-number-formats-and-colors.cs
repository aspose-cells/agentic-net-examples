// Title: Keep Custom Number Format & Cell Color When Refreshing an Aspose.Cells Pivot Table (C#)
// Description: Demonstrates how to enable PreserveFormatting on an Aspose.Cells PivotTable, apply a currency number format with a background fill, refresh the data, and save the workbook while retaining the custom styling.
// Keywords: Aspose.Cells PivotTable PreserveFormatting | C# custom number format pivot | pivot table cell background color | keep formatting after RefreshData | Aspose.Cells style data body range | Excel pivot formatting .NET
// Common Searches: Aspose.Cells keep pivot formatting after refresh | C# set PreserveFormatting for PivotTable | apply currency format to pivot table cells Aspose | preserve cell colors in Aspose.Cells pivot | how to format pivot data body range in .NET
// Developer Intent: Retain custom number formats and background colors on a PivotTable after calling RefreshData.
// Use Cases: Apply a currency format with a solid fill to the data area of a PivotTable and ensure the style survives data refreshes. | Create a sales report workbook where the PivotTable’s appearance remains unchanged after source data updates. | Programmatically generate Excel files with consistent PivotTable styling across multiple refresh cycles.
// AI Prompts: Show C# code that sets PreserveFormatting = true and formats a PivotTable data body range with a custom currency style using Aspose.Cells. | Explain how to refresh an Aspose.Cells PivotTable without losing custom number formats and cell colors. | Provide a step‑by‑step example of applying a background color and custom number format to a PivotTable and preserving them after RefreshData.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to enable PreserveFormatting on an Aspose.Cells PivotTable, apply a currency number format with a background fill, refresh the data, and save the workbook while retaining the custom styling.
class PreservePivotFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample source data for the pivot table
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B4"].PutValue(1500);

        // Add a pivot table based on the source range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot fields (row and data)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Ensure that formatting is kept when the pivot table is refreshed
        pivotTable.PreserveFormatting = true;

        // Create a style that includes a custom number format and a background color
        Style style = workbook.CreateStyle();
        style.Custom = "$#,##0.00";                 // Currency number format
        style.ForegroundColor = Color.LightBlue;    // Cell background color
        style.Pattern = BackgroundType.Solid;       // Apply the background color

        // Apply the style to the data body range of the pivot table
        CellArea dataArea = pivotTable.DataBodyRange;
        pivotTable.Format(dataArea, style);

        // Refresh the pivot table data and recalculate to reflect any changes
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the preserved formatting
        workbook.Save("PreserveFormattingPivot.xlsx", SaveFormat.Xlsx);
    }
}
