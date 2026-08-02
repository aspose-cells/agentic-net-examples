// Title: Aspose.Cells C# – Refresh Pivot Table After Header Formatting to Preserve Styles
// Description: Demonstrates how to create a workbook, add a pivot table, apply a bold white‑on‑dark‑blue header style, enable PreserveFormatting, and refresh the pivot so the custom header formatting remains in the saved XLSX file.
// Keywords: Aspose.Cells | C# | .NET | pivot table | RefreshPivotTables | PreserveFormatting | header style | Excel export | XLSX | PivotTable.Format | pivot refresh
// Common Searches: Aspose.Cells keep pivot header formatting after refresh | C# refresh pivot table preserve style | How to preserve custom pivot table formatting in Aspose.Cells | Refresh all pivot tables in worksheet Aspose.Cells .NET | PivotTable PreserveFormatting property example
// Developer Intent: Maintain custom header formatting when a pivot table is refreshed.
// Use Cases: Apply a bold, white‑on‑blue style to a pivot table header and ensure it survives data recalculation. | Programmatically refresh multiple pivot tables in a worksheet without losing user‑defined styles. | Generate Excel reports with styled pivot tables that stay consistent after updates.
// AI Prompts: Show C# code using Aspose.Cells to format a pivot table header and keep the style after refreshing. | Explain how PreserveFormatting works with RefreshPivotTables in Aspose.Cells for .NET. | Provide an example that creates a pivot table, styles its header, and refreshes it while preserving formatting.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add a pivot table, apply a bold white‑on‑dark‑blue header style, enable PreserveFormatting, and refresh the pivot so the custom header formatting remains in the saved XLSX file.
class RefreshPivotAfterHeaderFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Fruit");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Vegetable");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Fruit");
        sheet.Cells["B4"].PutValue(15);

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table (row and data areas)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Calculate the pivot table to populate it with data
        pivotTable.CalculateData();

        // Create a style for the header cell (bold font, white text on dark blue background)
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.White;
        headerStyle.ForegroundColor = Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;

        // Format the header cell of the pivot table.
        // In pivot coordinates, the first header cell is at row index 1, column index 0.
        pivotTable.Format(1, 0, headerStyle);

        // Ensure that formatting is preserved when the pivot table is refreshed
        pivotTable.PreserveFormatting = true;

        // Refresh all pivot tables in the worksheet so that the style persists
        sheet.RefreshPivotTables();

        // Save the workbook
        workbook.Save("PivotHeaderFormatted.xlsx", SaveFormat.Xlsx);
    }
}
