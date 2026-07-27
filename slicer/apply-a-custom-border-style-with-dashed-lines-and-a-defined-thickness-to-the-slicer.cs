using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerBorderDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Fruit");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["A4"].PutValue("Banana");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(15);
                worksheet.Cells["B4"].PutValue(20);

                // Add a pivot table based on the sample data
                int pivotIdx = worksheet.PivotTables.Add("A1:B4", "E3", "FruitPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIdx];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // Add "Fruit" as row field

                // Add a slicer linked to the pivot table
                int slicerIdx = worksheet.Slicers.Add(pivotTable, "A1", "FruitSlicer");
                Slicer slicer = worksheet.Slicers[slicerIdx];

                // Optional: set a built‑in style for the slicer (keeps default appearance)
                slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

                // Access the underlying shape of the slicer to customize its border
                Shape slicerShape = slicer.Shape;

                // Ensure the shape has a visible line (border)
                slicerShape.HasLine = true;

                // Set the line (border) weight and dash style
                slicerShape.Line.Weight = 2.0;               // 2 points thickness
                slicerShape.Line.DashStyle = MsoLineDashStyle.Dash;

                // If the LineFormat exposes a Color property, you can set it as follows:
                // slicerShape.Line.Color = Color.DarkBlue;

                // Save the workbook with the customized slicer
                string outputPath = "SlicerWithCustomBorder.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}