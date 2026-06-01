using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

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
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("Fruit");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Vegetable");
            cells["B3"].PutValue(80);
            cells["A4"].PutValue("Fruit");
            cells["B4"].PutValue(150);
            cells["A5"].PutValue("Vegetable");
            cells["B5"].PutValue(200);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D2", "PivotTable1");
            PivotTable pivot = worksheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table (Category field)
            // Correct parameter order: destination cell name first, then field name
            int slicerIndex = worksheet.Slicers.Add(pivot, "E2", "Category");
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Change the slicer layout direction to right‑to‑left
            slicer.Shape.TextDirection = TextDirectionType.RightToLeft;

            // Optionally set the whole worksheet to display right‑to‑left
            worksheet.DisplayRightToLeft = true;

            // Save the workbook
            workbook.Save("SlicerRTL.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}