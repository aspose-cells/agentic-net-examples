using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

class SlicerRtlDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Fruit";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Vegetable";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Grain";
            cells["B4"].Value = 150;

            // Create a pivot table based on the data range
            PivotTableCollection pivots = worksheet.PivotTables;
            int pivotIndex = pivots.Add("A1:B4", "D5", "PivotTable1");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the "Category" field of the pivot table
            // Note: The correct parameter order is (pivot, destination cell, base field name)
            SlicerCollection slicers = worksheet.Slicers;
            int slicerIndex = slicers.Add(pivot, "F2", "Category");
            Slicer slicer = slicers[slicerIndex];

            // Change the slicer layout direction to Right‑to‑Left (RTL)
            slicer.Shape.TextDirection = TextDirectionType.RightToLeft;

            // Optionally set the worksheet itself to display right‑to‑left
            worksheet.DisplayRightToLeft = true;

            // Save the workbook
            workbook.Save("SlicerRtlDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}