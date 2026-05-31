using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Food");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Transport");
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["A4"].PutValue("Utilities");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table based on the data range
        int ptIndex = sheet.PivotTables.Add("A1:B4", "D6", "MyPivot");
        PivotTable pivot = sheet.PivotTables[ptIndex];

        // Configure the pivot table fields
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Ensure pivot data is saved with the workbook
        pivot.SaveData = true;

        // Preserve formatting when the pivot table is refreshed
        pivot.PreserveFormatting = true;

        // Apply a built‑in pivot table style
        pivot.PivotTableStyleName = "PivotStyleMedium9";

        // Create a custom style for the data body
        Style customStyle = workbook.CreateStyle();
        customStyle.Font.Name = "Calibri";
        customStyle.Font.Size = 11;
        customStyle.Font.IsBold = true;
        customStyle.ForegroundColor = Color.LightYellow;
        customStyle.Pattern = BackgroundType.Solid;

        // Apply the custom style to the pivot table's data body range
        pivot.Format(pivot.DataBodyRange, customStyle);

        // Refresh and calculate the pivot table to reflect any changes
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the modified workbook as XLSX
        workbook.Save("ModifiedPivotWorkbook.xlsx", SaveFormat.Xlsx);
    }
}