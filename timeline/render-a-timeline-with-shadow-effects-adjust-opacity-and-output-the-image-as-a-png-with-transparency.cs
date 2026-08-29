// Title: Render an Excel timeline with a custom shadow and export it as a transparent PNG using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a pivot table, adds a linked timeline, customizes the timeline's shadow (angle, blur, size, distance, transparency) and saves the worksheet as a PNG with an alpha channel using Aspose.Cells. | Show how to configure ImageOrPrintOptions for PNG transparency and render a worksheet containing a timeline to an image file in C# with Aspose.Cells. | Demonstrate modifying the TimelineShape.ShadowEffect properties before exporting the timeline to a transparent PNG in a .NET application.
// Common Searches: how to export an Excel timeline as a transparent PNG with Aspose.Cells C# | Aspose.Cells timeline shadow effect customization example | C# render worksheet timeline to PNG with alpha channel | set timeline shape shadow transparency using Aspose.Cells .NET | image rendering options transparent PNG Aspose.Cells timeline
// Tags: timeline shadow customization Aspose.Cells | transparent PNG export Aspose.Cells | ImageOrPrintOptions PNG transparency .NET | TimelineShape shadow properties C# | pivot table linked timeline Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace TimelineRenderDemo
{
    // The example creates a workbook with sample data, builds a pivot table, adds a timeline linked to the date field, customizes the timeline's shadow (angle, blur, size, distance, transparency), sets ImageOrPrintOptions for PNG with a transparent background, and renders the worksheet to a PNG file that preserves the shadow and alpha channel.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (fruit, date, amount)
            cells[0, 0].Value = "Fruit";
            cells[0, 1].Value = "Date";
            cells[0, 2].Value = "Amount";

            cells[1, 0].Value = "Apple";
            cells[1, 1].Value = new DateTime(2021, 1, 10);
            cells[1, 2].Value = 120;

            cells[2, 0].Value = "Banana";
            cells[2, 1].Value = new DateTime(2021, 2, 15);
            cells[2, 2].Value = 150;

            cells[3, 0].Value = "Cherry";
            cells[3, 1].Value = new DateTime(2021, 3, 20);
            cells[3, 2].Value = 180;

            // Apply date style to the date column
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";
            cells[1, 1].SetStyle(dateStyle);
            cells[2, 1].SetStyle(dateStyle);
            cells[3, 1].SetStyle(dateStyle);

            // Create a PivotTable based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("=Sheet1!A1:C4", "E5", "FruitPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable's date field
            sheet.Timelines.Add(pivot, 15, 0, "Date");
            Timeline timeline = sheet.Timelines[0];

            // Configure Timeline properties (optional)
            timeline.Caption = "Sales Timeline";
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;
            timeline.ShowSelectionLabel = true;
            timeline.ShowTimeLevel = true;

            // Access the underlying shape to apply shadow effect
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
                Transparent = true   // enable transparent background
            };

            // Render the worksheet (which contains the timeline) to a PNG image
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            // Render the first (and only) page to a file
            renderer.ToImage(0, "TimelineWithShadow.png");

            Console.WriteLine("Timeline rendered to PNG with shadow effect and transparency.");
        }
    }
}
