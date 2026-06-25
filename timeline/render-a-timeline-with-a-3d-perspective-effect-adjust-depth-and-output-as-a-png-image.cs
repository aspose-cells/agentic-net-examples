using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class RenderTimeline3D
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a style for date cells (short date format)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // short date format

            // Populate worksheet with sample date/value data
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");

            // Row 2
            cells["A2"].PutValue(new DateTime(2023, 1, 1));
            cells["A2"].SetStyle(dateStyle);
            cells["B2"].PutValue(100);

            // Row 3
            cells["A3"].PutValue(new DateTime(2023, 2, 1));
            cells["A3"].SetStyle(dateStyle);
            cells["B3"].PutValue(150);

            // Row 4
            cells["A4"].PutValue(new DateTime(2023, 3, 1));
            cells["A4"].SetStyle(dateStyle);
            cells["B4"].PutValue(200);

            // Create a PivotTable based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable (row 10, column 5)
            int timelineIndex = sheet.Timelines.Add(pivot, 10, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Adjust the size of the Timeline shape
            timeline.Shape.Width = 400;   // width in pixels
            timeline.Shape.Height = 120;  // height in pixels

            // Apply 3‑D perspective effect via ThreeDFormat
            ThreeDFormat threeD = timeline.Shape.ThreeDFormat;
            threeD.Perspective = 45;      // perspective angle (0‑120 degrees)
            threeD.ExtrusionHeight = 20;  // give the shape depth
            threeD.Z = 10;                // distance from the ground

            // Render the worksheet (which contains the Timeline) to a PNG image
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png
            };
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToImage(0, "Timeline3D.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}