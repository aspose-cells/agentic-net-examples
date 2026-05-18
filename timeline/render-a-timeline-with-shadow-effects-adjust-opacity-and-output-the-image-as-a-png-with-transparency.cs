using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;

namespace AsposeCellsTimelineRender
{
    public class RenderTimelineWithShadow
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate worksheet with sample data (including a date column)
                cells[0, 0].Value = "Fruit";
                cells[0, 1].Value = "Date";
                cells[0, 2].Value = "Amount";

                cells[1, 0].Value = "Apple";
                cells[1, 1].Value = new DateTime(2021, 2, 5);
                cells[1, 2].Value = 120;

                cells[2, 0].Value = "Orange";
                cells[2, 1].Value = new DateTime(2022, 3, 8);
                cells[2, 2].Value = 150;

                cells[3, 0].Value = "Banana";
                cells[3, 1].Value = new DateTime(2023, 4, 10);
                cells[3, 2].Value = 180;

                // Create a PivotTable based on the data
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("=Sheet1!A1:C4", "E5", "FruitPivot");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Add the Date field as a Page (filter) field – required for Timeline creation
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");

                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline linked to the PivotTable's date field
                sheet.Timelines.Add(pivot, 15, 0, "Date");
                Timeline timeline = sheet.Timelines[0];

                // Configure timeline appearance
                timeline.Caption = "Sales Timeline";
                timeline.ShowHeader = true;
                timeline.ShowHorizontalScrollbar = true;

                // Apply shadow effect to the underlying shape
                TimelineShape timelineShape = timeline.Shape;
                ShadowEffect shadow = timelineShape.ShadowEffect;
                shadow.Angle = 135;          // direction of the shadow
                shadow.Blur = 20;            // blur radius
                shadow.Size = 1.0;           // size multiplier
                shadow.Distance = 10;        // distance from the shape
                shadow.Transparency = 0.4;   // 40% transparent shadow

                // Set image rendering options: PNG format with transparent background
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    Transparent = true
                };

                // Render the worksheet (including the timeline) to a PNG image
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                string outputPath = "TimelineWithShadow.png";
                renderer.ToImage(0, outputPath);

                Console.WriteLine($"Timeline rendered to image with shadow and transparency: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}