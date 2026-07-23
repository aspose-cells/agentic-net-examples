// Title: Aspose.Cells for .NET – Render a PivotTable Timeline with Shadow and Transparent PNG
// Description: This C# example creates a workbook, fills it with fruit, date, and amount data, builds a PivotTable, adds a linked Timeline, applies a semi‑transparent shadow (angle, blur, size, distance) to the TimelineShape, and exports the shape as a PNG with a transparent background while also saving the workbook.
// Keywords: Aspose.Cells | C# timeline rendering | PivotTable timeline | timeline shadow effect | transparent PNG export | timeline shape to image | shadow transparency Aspose.Cells | image rendering .NET | export timeline as PNG | Aspose.Cells timeline example
// Common Searches: How to add a shadow to an Aspose.Cells timeline | Export Aspose.Cells timeline to PNG with transparency | Set shadow angle and blur for timeline shape C# | Render timeline shape to image using Aspose.Cells | Aspose.Cells timeline transparent background
// Developer Intent: Generate a timeline linked to a PivotTable, style its shadow, and save it as a transparent PNG image.
// Use Cases: Create presentation‑ready timeline graphics with a subtle shadow effect. | Produce PNG assets for web dashboards where the background must stay transparent. | Automate batch reporting that exports timeline visuals for email or documentation.
// AI Prompts: Show me how to change the shadow transparency of a timeline shape in Aspose.Cells and export it as a transparent PNG. | Provide C# code to render multiple PivotTable timelines into separate PNG files with shadow styling using Aspose.Cells. | Explain how to configure shadow angle, blur, size, and distance for a timeline shape while preserving PNG transparency.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineRender
{
    // This C# example creates a workbook, fills it with fruit, date, and amount data, builds a PivotTable, adds a linked Timeline, applies a semi‑transparent shadow (angle, blur, size, distance) to the TimelineShape, and exports the shape as a PNG with a transparent background while also saving the workbook.
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
            cells[0, 1].Value = "date";
            cells[0, 2].Value = "amount";

            string[] fruits = { "grape", "blueberry", "kiwi", "cherry" };
            DateTime[] dates = {
                new DateTime(2021, 2, 5),
                new DateTime(2022, 3, 8),
                new DateTime(2023, 4, 10),
                new DateTime(2024, 5, 16)
            };
            int[] amounts = { 50, 60, 70, 80 };

            // Apply date style
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";

            for (int i = 0; i < fruits.Length; i++)
            {
                cells[i + 1, 0].Value = fruits[i];
                cells[i + 1, 1].Value = dates[i];
                cells[i + 1, 1].SetStyle(dateStyle);
                cells[i + 1, 2].Value = amounts[i];
            }

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

            // Position and size the timeline (optional)
            timeline.LeftPixel = 100;
            timeline.TopPixel = 50;
            timeline.WidthPixel = 400;
            timeline.HeightPixel = 120;

            // Access the underlying shape of the timeline
            TimelineShape timelineShape = timeline.Shape;

            // Configure shadow effect: set transparency (0.0 = opaque, 1.0 = clear)
            ShadowEffect shadow = timelineShape.ShadowEffect;
            shadow.Transparency = 0.4; // 40% transparent shadow
            shadow.Angle = 135;
            shadow.Blur = 20;
            shadow.Size = 1.0;
            shadow.Distance = 10;

            // Prepare image options: PNG format with transparent background
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                Transparent = true // make background transparent
            };

            // Render the timeline shape to a PNG file
            string outputImagePath = "TimelineWithShadow.png";
            timelineShape.ToImage(outputImagePath, imgOptions);

            // Optionally save the workbook for reference
            workbook.Save("TimelineWorkbook.xlsx");

            Console.WriteLine($"Timeline rendered to image: {outputImagePath}");
        }
    }
}
