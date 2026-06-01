using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTimeline3D
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate worksheet with sample data (Date + Value)
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");

                sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
                sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
                sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 1));

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(210);

                // -------------------------------------------------
                // Create a PivotTable that will be the data source for the Timeline
                // -------------------------------------------------
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // Add a Timeline linked to the PivotTable
                // -------------------------------------------------
                // Place the Timeline starting at row 10, column 2 (cell B10)
                int timelineIndex = sheet.Timelines.Add(pivot, 9, 1, "Date");
                var timeline = sheet.Timelines[timelineIndex]; // Timeline object

                // -------------------------------------------------
                // Apply 3‑D perspective effect to the Timeline shape
                // -------------------------------------------------
                var shape = timeline.Shape;                 // TimelineShape
                var threeD = shape.ThreeDFormat;            // ThreeDFormat

                // Set perspective angle (0‑120 degrees). 45 gives a moderate 3‑D view.
                threeD.Perspective = 45;

                // Adjust depth (distance from the ground) to enhance the 3‑D effect.
                threeD.Z = 30;               // moves the shape away from the sheet
                threeD.ExtrusionHeight = 10; // gives the shape some thickness

                // Optional: rotate for better visual appearance
                threeD.RotationX = 20;
                threeD.RotationY = 30;

                // -------------------------------------------------
                // Render the worksheet (which now contains the Timeline) to a PNG image
                // -------------------------------------------------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true
                };

                SheetRender renderer = new SheetRender(sheet, renderOptions);
                // Render the first (and only) page to a file named "Timeline3D.png"
                renderer.ToImage(0, "Timeline3D.png");
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}