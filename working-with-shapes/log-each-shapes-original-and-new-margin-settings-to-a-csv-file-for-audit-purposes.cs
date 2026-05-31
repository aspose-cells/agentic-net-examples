using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsShapeMarginAudit
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate margin logging
            Shape shape = sheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);
            shape.Name = "DemoRectangle";

            // Set some initial margin values
            ShapeTextAlignment alignment = shape.TextBody.TextAlignment;
            alignment.LeftMarginPt = 5.0;
            alignment.RightMarginPt = 5.0;
            alignment.TopMarginPt = 2.0;
            alignment.BottomMarginPt = 2.0;
            alignment.IsAutoMargin = false;

            // Prepare CSV file
            string csvPath = "ShapeMarginsAudit.csv";
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("ShapeName,OriginalLeft,OriginalRight,OriginalTop,OriginalBottom,NewLeft,NewRight,NewTop,NewBottom");

                // Iterate through all shapes in the worksheet
                foreach (Shape shp in sheet.Shapes)
                {
                    // Capture original margins
                    ShapeTextAlignment originalAlignment = shp.TextBody.TextAlignment;
                    double origLeft = originalAlignment.LeftMarginPt;
                    double origRight = originalAlignment.RightMarginPt;
                    double origTop = originalAlignment.TopMarginPt;
                    double origBottom = originalAlignment.BottomMarginPt;

                    // Modify margins (example: increase each by 1 point)
                    originalAlignment.LeftMarginPt = origLeft + 1.0;
                    originalAlignment.RightMarginPt = origRight + 1.0;
                    originalAlignment.TopMarginPt = origTop + 1.0;
                    originalAlignment.BottomMarginPt = origBottom + 1.0;

                    // Capture new margins
                    double newLeft = originalAlignment.LeftMarginPt;
                    double newRight = originalAlignment.RightMarginPt;
                    double newTop = originalAlignment.TopMarginPt;
                    double newBottom = originalAlignment.BottomMarginPt;

                    // Write a line to CSV
                    writer.WriteLine($"{shp.Name},{origLeft},{origRight},{origTop},{origBottom},{newLeft},{newRight},{newTop},{newBottom}");
                }
            }

            // Save the workbook with the shape
            workbook.Save("ShapeMarginAuditWorkbook.xlsx");
        }
    }
}