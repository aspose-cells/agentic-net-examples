using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeReport
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Add some sample shapes to the first worksheet
        // -------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Shapes.AddRectangle(2, 2, 2, 2, 100, 200);               // Rectangle
        sheet.Shapes.AddLine(5, 5, 5, 5, 150, 1);                     // Line

        // Auto shape with adjustment guides
        Shape autoShape = sheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 8, 8, 0, 0, 200, 100);
        autoShape.Geometry.ShapeAdjustValues.Add("Roundness", 0.2);
        autoShape.Geometry.ShapeAdjustValues.Add("Angle", 45);

        // -------------------------------------------------
        // Create a summary worksheet
        // -------------------------------------------------
        int summaryIndex = workbook.Worksheets.Add();
        Worksheet summary = workbook.Worksheets[summaryIndex];
        summary.Name = "ShapeSummary";

        // Header row
        summary.Cells["A1"].PutValue("Worksheet");
        summary.Cells["B1"].PutValue("Shape Index");
        summary.Cells["C1"].PutValue("Type");
        summary.Cells["D1"].PutValue("UpperLeftRow");
        summary.Cells["E1"].PutValue("UpperLeftColumn");
        summary.Cells["F1"].PutValue("LowerRightRow");
        summary.Cells["G1"].PutValue("LowerRightColumn");
        summary.Cells["H1"].PutValue("Adjust Values");

        int currentRow = 1; // zero‑based index for the next data row

        // -------------------------------------------------
        // Enumerate all worksheets (except the summary sheet)
        // -------------------------------------------------
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws == summary) continue; // skip the summary sheet itself

            ShapeCollection shapes = ws.Shapes;
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];

                // Basic shape information
                summary.Cells[currentRow, 0].PutValue(ws.Name);
                summary.Cells[currentRow, 1].PutValue(i);
                summary.Cells[currentRow, 2].PutValue(shape.Type.ToString());
                summary.Cells[currentRow, 3].PutValue(shape.UpperLeftRow);
                summary.Cells[currentRow, 4].PutValue(shape.UpperLeftColumn);
                summary.Cells[currentRow, 5].PutValue(shape.LowerRightRow);
                summary.Cells[currentRow, 6].PutValue(shape.LowerRightColumn);

                // Collect adjustment guide values, if any
                string adjustInfo = "";
                if (shape.Geometry != null && shape.Geometry.ShapeAdjustValues != null)
                {
                    foreach (ShapeGuide guide in shape.Geometry.ShapeAdjustValues)
                    {
                        adjustInfo += $"{guide.Value}; ";
                    }
                }
                summary.Cells[currentRow, 7].PutValue(adjustInfo.Trim());

                currentRow++;
            }
        }

        // Auto‑fit columns for better readability
        summary.AutoFitColumns();

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("ShapeReport.xlsx");
    }
}