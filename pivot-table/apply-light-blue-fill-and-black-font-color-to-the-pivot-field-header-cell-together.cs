using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHeaderFormatting
{
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

            // Add a pivot table (source range A1:B4, placed at D3)
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate the pivot data
            pivotTable.CalculateData();

            // Create a style: light blue fill and black font color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;      // Fill color
            headerStyle.Pattern = BackgroundType.Solid;        // Apply fill
            headerStyle.Font.Color = Color.Black;              // Font color
            headerStyle.Font.IsBold = true;                    // Optional: make header bold

            // Format the pivot field header cell.
            // The Format(row, column, style) method uses pivot‑table‑relative coordinates.
            // Row 2, Column 0 corresponds to the first field header cell.
            pivotTable.Format(2, 0, headerStyle);

            // Save the workbook
            workbook.Save("PivotHeaderFormatted.xlsx");
        }
    }
}