using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
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

        // Configure pivot fields
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Apply pivot table properties
        pivot.SaveData = true;                     // Ensure pivot data is saved with the file
        pivot.PreserveFormatting = true;           // Preserve formatting on refresh
        pivot.PivotTableStyleName = "PivotStyleLight14"; // Apply a built‑in style

        // Create a custom style for the data area
        Style customStyle = workbook.CreateStyle();
        customStyle.Font.Name = "Calibri";
        customStyle.Font.Size = 11;
        customStyle.Font.IsBold = true;
        customStyle.ForegroundColor = Color.LightGreen;
        customStyle.Pattern = BackgroundType.Solid;

        // Apply the custom style to the pivot table's data body range
        CellArea dataArea = pivot.DataBodyRange;
        pivot.Format(dataArea, customStyle);

        // Refresh and calculate the pivot table to reflect any changes
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the modified workbook as XLSX
        workbook.Save("ModifiedPivotWorkbook.xlsx", SaveFormat.Xlsx);
    }
}