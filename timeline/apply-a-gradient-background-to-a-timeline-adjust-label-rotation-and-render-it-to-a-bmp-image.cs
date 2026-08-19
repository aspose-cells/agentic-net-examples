// Title: C# – Apply Gradient Fill & Rotate Labels on an Aspose.Cells Timeline, Export to BMP
// Description: Creates a workbook, builds a pivot table, adds a timeline linked to the date field, applies a horizontal two‑color gradient to the timeline shape, rotates the labels 45°, and renders the sheet as a BMP image using Aspose.Cells for .NET.
// Keywords: Aspose.Cells timeline gradient | C# timeline label rotation | export timeline to BMP | pivot table timeline Aspose.Cells | two‑color gradient fill shape | SheetRender BMP Aspose.Cells | TimelineShape FillType Gradient | Aspose.Cells example C#
// Common Searches: how to add gradient fill to a timeline in Aspose.Cells .NET | rotate timeline labels Aspose.Cells C# | export worksheet with timeline as BMP image | create pivot table and timeline programmatically Aspose.Cells | set horizontal two‑color gradient on timeline shape
// Developer Intent: Generate a timeline tied to a pivot table, style it with a two‑color gradient, tilt the labels for better readability, and save the resulting sheet as a BMP file.
// Use Cases: Design a sales‑trend visual with a blue gradient for inclusion in presentations. | Produce a printable report where rotated month labels prevent overlap. | Create a web‑ready BMP snapshot of a styled timeline for dashboards.
// AI Prompts: Show how to change the gradient colors and direction of a timeline shape in Aspose.Cells for .NET. | Provide code to set a different label rotation angle and render the sheet as PNG instead of BMP. | Explain how to add multiple timelines to a worksheet and export each to separate image files.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, builds a pivot table, adds a timeline linked to the date field, applies a horizontal two‑color gradient to the timeline shape, rotates the labels 45°, and renders the sheet as a BMP image using Aspose.Cells for .NET.
class TimelineGradientRender
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table (date and value)
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue(new DateTime(2023, 1, 1));
            cells["B2"].PutValue(100);
            cells["A3"].PutValue(new DateTime(2023, 2, 1));
            cells["B3"].PutValue(200);
            cells["A4"].PutValue(new DateTime(2023, 3, 1));
            cells["B4"].PutValue(300);

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Place the date field in the column area (required for timeline)
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            // Place the value field in the data area
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table using the date field
            int timelineIdx = sheet.Timelines.Add(pivot, 10, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];

            // Access the underlying shape of the timeline
            TimelineShape shape = timeline.Shape;

            // Apply a two‑color gradient background to the timeline shape
            shape.Fill.FillType = FillType.Gradient;
            shape.Fill.SetTwoColorGradient(Color.LightBlue, Color.DarkBlue, GradientStyleType.Horizontal, 1);

            // Rotate the timeline labels (shape rotation) by 45 degrees
            shape.RotationAngle = 45;

            // Render the worksheet (which contains the timeline) to a BMP image
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Bmp,
                OnePagePerSheet = true
            };
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "TimelineOutput.bmp");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
