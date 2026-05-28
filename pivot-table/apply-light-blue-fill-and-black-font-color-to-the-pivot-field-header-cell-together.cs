using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Bike");
        worksheet.Cells["B2"].PutValue(5000);
        worksheet.Cells["A3"].PutValue("Car");
        worksheet.Cells["B3"].PutValue(12000);
        worksheet.Cells["A4"].PutValue("Truck");
        worksheet.Cells["B4"].PutValue(8000);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add fields to the pivot table (row and data)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Calculate the pivot table data
        pivotTable.CalculateData();

        // Create a style with light blue fill and black font color
        Style headerStyle = workbook.CreateStyle();
        headerStyle.ForegroundColor = Color.LightBlue;      // Light blue fill
        headerStyle.Pattern = BackgroundType.Solid;        // Apply fill pattern
        headerStyle.Font.Color = Color.Black;              // Black font color

        // Apply the style to the pivot field header cell.
        // According to Aspose.Cells examples, the header cell is at row index 2, column index 0
        // in pivot table coordinates.
        pivotTable.Format(2, 0, headerStyle);

        // Save the workbook
        workbook.Save("PivotHeaderFormatted.xlsx");
    }
}