// Title: Render a 3‑D Timeline from a Pivot Table to PNG with Aspose.Cells (C#)
// Description: This example creates a workbook, fills it with date/value data, builds a pivot table, adds a linked timeline, applies a 3‑D perspective (perspective angle, extrusion height, X/Y rotation) to the timeline shape, and renders the worksheet as a PNG image using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | timeline | 3D perspective | pivot table | PNG export | ThreeDFormat | Excel visualization | .NET | render worksheet as image
// Common Searches: Aspose.Cells add 3D effect to timeline | C# export timeline as PNG | How to render Excel timeline image with Aspose.Cells | Set perspective angle for timeline shape .NET | Create pivot table timeline and save as image
// Developer Intent: Generate a pivot‑driven timeline, style it with 3‑D depth, and save the result as a PNG file using Aspose.Cells for .NET.
// Use Cases: Produce a sales‑by‑month dashboard thumbnail with a depth‑styled timeline for PDF reports. | Automate weekly status emails by rendering project‑milestone timelines as PNG images. | Batch‑process workbooks to add consistently styled 3‑D timelines and publish them on a website.
// AI Prompts: Write C# code that uses Aspose.Cells to create a pivot table, add a linked timeline, apply ThreeDFormat (perspective, extrusion height, rotation X/Y), and export the sheet as a PNG image. | Explain how to configure the ThreeDFormat properties of a timeline shape in Aspose.Cells to achieve a specific 3‑D perspective. | Show how to render only the timeline shape to a PNG file without the surrounding worksheet cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This example creates a workbook, fills it with date/value data, builds a pivot table, adds a linked timeline, applies a 3‑D perspective (perspective angle, extrusion height, X/Y rotation) to the timeline shape, and renders the worksheet as a PNG image using Aspose.Cells for .NET.
class Timeline3DPng
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample date/value data
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(200);

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Place the Date field in the column area (required for Timeline)
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            // Place the Value field in the data area
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table (placed at row 10, column 0)
            int timelineIndex = sheet.Timelines.Add(pivot, 10, 0, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Adjust size and position of the timeline shape
            timeline.Shape.Width = 400;   // width in pixels
            timeline.Shape.Height = 80;   // height in pixels
            timeline.Shape.Left = 0;
            timeline.Shape.Top = 0;

            // Apply 3‑D perspective effect to the timeline shape
            ThreeDFormat threeD = timeline.Shape.ThreeDFormat;
            threeD.Perspective = 45;      // perspective angle (0‑120 degrees)
            threeD.ExtrusionHeight = 10;  // give the shape some depth
            threeD.RotationX = 20;        // rotate around X‑axis
            threeD.RotationY = 30;        // rotate around Y‑axis

            // Render the worksheet (which now contains the 3‑D timeline) to a PNG image
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png
            };
            SheetRender renderer = new SheetRender(sheet, renderOptions);
            renderer.ToImage(0, "Timeline3D.png"); // page index 0, output file name
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
