using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the pivot table
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Sales";
        sheet.Cells["A2"].Value = "Bike";
        sheet.Cells["B2"].Value = 5000;
        sheet.Cells["A3"].Value = "Car";
        sheet.Cells["B3"].Value = 12000;
        sheet.Cells["A4"].Value = "Truck";
        sheet.Cells["B4"].Value = 8000;

        // Create a pivot table
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Populate the pivot table
        pivotTable.CalculateData();

        // Assign the built‑in style PivotTableStyleMedium9
        pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

        // Save the workbook (lifecycle save)
        workbook.Save("PivotTableStyleMedium9.xlsx");
    }
}