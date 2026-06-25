using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineRender
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (fruit, date, amount)
            cells[0, 0].Value = "fruit";
            cells[1, 0].Value = "grape";
            cells[2, 0].Value = "blueberry";
            cells[3, 0].Value = "kiwi";
            cells[4, 0].Value = "cherry";

            // Create a date style
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";

            cells[0, 1].Value = "date";
            cells[1, 1].Value = new DateTime(2021, 2, 5);
            cells[2, 1].Value = new DateTime(2022, 3, 8);
            cells[3, 1].Value = new DateTime(2023, 4, 10);
            cells[4, 1].Value = new DateTime(2024, 5, 16);
            cells[1, 1].SetStyle(dateStyle);
            cells[2, 1].SetStyle(dateStyle);
            cells[3, 1].SetStyle(dateStyle);
            cells[4, 1].SetStyle(dateStyle);

            cells[0, 2].Value = "amount";
            cells[1, 2].Value = 50;
            cells[2, 2].Value = 60;
            cells[3, 2].Value = 70;
            cells[4, 2].Value = 80;

            // Add a PivotTable based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("=Sheet1!A1:C5", "A12", "FruitPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "date");
            pivot.AddFieldToArea(PivotFieldType.Data, "amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable's date field
            sheet.Timelines.Add(pivot, 10, 5, "date");
            Timeline timeline = sheet.Timelines[0];

            // Optional: adjust timeline visual properties
            timeline.Caption = "Sales Timeline";
            timeline.LeftPixel = 100;
            timeline.TopPixel = 50;
            timeline.WidthPixel = 400;
            timeline.HeightPixel = 120;

            // Access the underlying shape (TimelineShape) to apply shadow effect
            Shape timelineShape = timeline.Shape;

            // Configure shadow effect: set transparency (0.0 = opaque, 1.0 = fully transparent)
            ShadowEffect shadow = timelineShape.ShadowEffect;
            shadow.Transparency = 0.4; // 40% transparent shadow
            shadow.Angle = 135;
            shadow.Blur = 20;
            shadow.Size = 1.5;
            shadow.Distance = 10;

            // Prepare image options: PNG format with transparent background
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                Transparent = true // enable transparent background
            };

            // Render the timeline shape to a PNG image file
            string outputPath = "TimelineImage.png";
            timelineShape.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Timeline rendered to image: {outputPath}");
        }
    }
}