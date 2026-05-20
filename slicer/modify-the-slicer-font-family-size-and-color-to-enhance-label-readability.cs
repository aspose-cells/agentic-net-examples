using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerFontDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["B3"].Value = 15;
            sheet.Cells["B4"].Value = 20;

            // Add a pivot table based on the sample data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Row field: Fruit
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Data field: Quantity
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "G1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Modify slicer font for better readability
            // Access the underlying shape of the slicer
            SlicerShape slicerShape = slicer.Shape as SlicerShape;
            if (slicerShape != null)
            {
                // Set desired font family, size, and color
                slicerShape.Font.Name = "Calibri";
                slicerShape.Font.Size = 12;
                slicerShape.Font.Color = Color.DarkBlue;
                // Optionally make the font bold
                slicerShape.Font.IsBold = true;
            }

            // Save the workbook with the customized slicer
            workbook.Save("SlicerFontDemo.xlsx");
        }
    }
}