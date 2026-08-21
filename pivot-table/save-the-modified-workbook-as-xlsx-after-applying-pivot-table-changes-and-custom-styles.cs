// Title: Aspose.Cells for .NET – Apply a Custom Pivot Table Style, Preserve Formatting, and Export to XLSX
// Description: C# example that creates a workbook, inserts sample data, builds a pivot table, enables SaveData and PreserveFormatting, applies the built‑in "PivotStyleMedium9" plus a custom yellow‑background style to the data body, refreshes the pivot, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells pivot table C# | custom pivot style Aspose.Cells | preserve pivot formatting .NET | save workbook as XLSX Aspose | RefreshPivotTables example | PivotTable.SaveData property | PivotTable.PreserveFormatting
// Common Searches: how to style pivot table data body in Aspose.Cells C# | save pivot table workbook as xlsx using Aspose.Cells | preserve pivot table formatting after refresh .NET | apply custom background color to Aspose.Cells pivot table | Aspose.Cells example for pivot table export
// Developer Intent: Add a styled pivot table, keep its formatting when refreshed, and generate an XLSX file with Aspose.Cells for .NET.
// Use Cases: Generate a sales summary workbook with a colored pivot table that retains its look after data updates. | Create a financial report where the pivot table must be exported to XLSX with corporate branding applied. | Automate the production of client‑ready spreadsheets that include a pivot table styled with custom fonts and background colors.
// AI Prompts: Show C# code using Aspose.Cells to create a pivot table, set SaveData and PreserveFormatting, apply a custom style to the data body, and save as XLSX. | Provide an Aspose.Cells .NET snippet that formats a pivot table with PivotStyleMedium9 and a yellow background, then refreshes and exports the workbook. | Explain step‑by‑step how to preserve pivot table formatting after a refresh and export the result to an XLSX file with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that creates a workbook, inserts sample data, builds a pivot table, enables SaveData and PreserveFormatting, applies the built‑in "PivotStyleMedium9" plus a custom yellow‑background style to the data body, refreshes the pivot, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Food");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Travel");
        sheet.Cells["B3"].PutValue(300);
        sheet.Cells["A4"].PutValue("Food");
        sheet.Cells["B4"].PutValue(80);
        sheet.Cells["A5"].PutValue("Travel");
        sheet.Cells["B5"].PutValue(150);

        // Add a pivot table based on the data range
        int ptIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[ptIndex];

        // Configure pivot fields: Category as row, Amount as data
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Ensure pivot data is saved with the workbook
        pivot.SaveData = true;

        // Preserve formatting when the pivot table is refreshed
        pivot.PreserveFormatting = true;

        // Apply a built‑in pivot table style
        pivot.PivotTableStyleName = "PivotStyleMedium9";

        // Create a custom style for the data body range
        Style customStyle = workbook.CreateStyle();
        customStyle.Font.Name = "Calibri";
        customStyle.Font.Size = 11;
        customStyle.Font.IsBold = true;
        customStyle.ForegroundColor = Color.LightYellow;
        customStyle.Pattern = BackgroundType.Solid;

        // Apply the custom style to the pivot table's data body range
        pivot.Format(pivot.DataBodyRange, customStyle);

        // Refresh all pivot tables to reflect any changes
        workbook.Worksheets.RefreshPivotTables();

        // Save the modified workbook as XLSX using the provided Save method
        workbook.Save("ModifiedPivotWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
