// Title: C# – Render a 3‑D Perspective Timeline from a Pivot Table to PNG with Aspose.Cells
// Description: Creates a workbook, adds date/value rows, builds a pivot table, inserts a linked timeline, applies a 45° perspective and 20‑point extrusion depth, adjusts the shape size, and renders the worksheet as a PNG image using Aspose.Cells for .NET.
// Keywords: Aspose.Cells timeline 3D | C# timeline PNG export | pivot table timeline Aspose | ThreeDFormat perspective Aspose.Cells | extrusion height timeline shape | render worksheet to PNG | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add 3D perspective to timeline | Export timeline as PNG in C# | Set extrusion height for timeline shape Aspose.Cells | Create timeline linked to pivot table .NET | Render worksheet image with Aspose.Cells
// Developer Intent: Generate a timeline with a 3‑D perspective from pivot data and save it as a PNG file using Aspose.Cells for .NET.
// Use Cases: Produce high‑impact 3‑D timeline graphics for sales or project reports. | Automate batch creation of timeline images for multiple datasets. | Embed static timeline PNGs in web dashboards or PDF documents where interactivity is not needed.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot‑linked timeline, apply a 45° perspective and 20‑point extrusion, then render it to PNG. | Show how to resize and reposition a timeline shape before exporting it as an image. | Give troubleshooting steps when a timeline does not appear in the rendered PNG using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, adds date/value rows, builds a pivot table, inserts a linked timeline, applies a 45° perspective and 20‑point extrusion depth, adjusts the shape size, and renders the worksheet as a PNG image using Aspose.Cells for .NET.
class Timeline3DPng
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample date and value data
            sheet.Cells["A1"].Value = "Date";
            sheet.Cells["B1"].Value = "Value";
            sheet.Cells["A2"].Value = new DateTime(2023, 1, 1);
            sheet.Cells["A3"].Value = new DateTime(2023, 2, 1);
            sheet.Cells["A4"].Value = new DateTime(2023, 3, 1);
            sheet.Cells["B2"].Value = 100;
            sheet.Cells["B3"].Value = 200;
            sheet.Cells["B4"].Value = 300;

            // Create a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table.
            // Use the field index (0) instead of the field name to avoid name‑lookup issues.
            int timelineIndex = sheet.Timelines.Add(pivot, 10, 5, 0);
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Access the underlying shape of the timeline
            TimelineShape timelineShape = timeline.Shape;

            // Apply 3‑D perspective and depth (extrusion height) to the timeline shape
            timelineShape.ThreeDFormat.Perspective = 45;      // Perspective angle (0‑120 degrees)
            timelineShape.ThreeDFormat.ExtrusionHeight = 20; // Depth of the 3‑D effect

            // Optionally adjust size and position
            timelineShape.Width = 400;
            timelineShape.Height = 80;
            timelineShape.Left = 50;
            timelineShape.Top = 30;

            // Render the worksheet (which contains the timeline) to a PNG image
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png
            };
            SheetRender sheetRender = new SheetRender(sheet, renderOptions);
            sheetRender.ToImage(0, "Timeline3D.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
