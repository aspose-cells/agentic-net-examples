// Title: Add a horizontal two‑color gradient to an Aspose.Cells timeline, rotate its caption, and export the worksheet as a BMP image using C#
// AI Prompts: Generate C# code that creates a pivot table, attaches a timeline, sets a horizontal LightBlue‑to‑DarkBlue two‑color gradient fill on the timeline shape, rotates the caption 45°, and saves the sheet as a BMP file with Aspose.Cells. | Show how to apply a two‑color gradient background to a TimelineShape, adjust its rotation angle, and render the containing worksheet to a BMP image while also saving the workbook as XLSX in .NET.
// Common Searches: c# aspose.cells timeline gradient background example | how to set caption angle for Aspose.Cells timeline | export worksheet with timeline to BMP image using Aspose.Cells .NET | link timeline to pivot table and apply gradient fill in Aspose.Cells | two‑color gradient fill on Aspose.Cells timeline shape tutorial
// Tags: timeline shape gradient Aspose.Cells | timeline caption rotation Aspose.Cells | export worksheet as BMP Aspose.Cells | pivot table timeline linkage Aspose.Cells | two‑color gradient timeline Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using System;
using System.Drawing;

// The example creates a workbook, fills it with sample data, builds a pivot table, adds a timeline linked to the pivot, applies a horizontal LightBlue‑to‑DarkBlue two‑color gradient to the timeline shape, rotates the caption 45°, renders the worksheet to a BMP file, and saves the workbook as XLSX.
class TimelineGradientExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");
            DateTime start = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                cells[1 + i, 0].PutValue(start.AddMonths(i));
                cells[1 + i, 1].PutValue(100 + i * 50);
            }

            // Create a pivot table based on the data
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:B6", "D1", "Pivot1");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table.
            // Use the field index (0) for the "Date" field to avoid name‑lookup issues.
            int timelineIdx = sheet.Timelines.Add(pivot, 10, 0, 0);
            Timeline timeline = sheet.Timelines[timelineIdx];

            // Apply a two‑color gradient background to the timeline shape
            TimelineShape shape = timeline.Shape;
            shape.Fill.FillType = FillType.Gradient;
            shape.Fill.SetTwoColorGradient(Color.LightBlue, Color.DarkBlue, GradientStyleType.Horizontal, 1);

            // Rotate the timeline (caption) for better label orientation
            shape.RotationAngle = 45;

            // Render the worksheet (which contains the timeline) to a BMP image
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Bmp
            };
            SheetRender render = new SheetRender(sheet, options);
            render.ToImage(0, "TimelineOutput.bmp");

            // Save the workbook (optional)
            workbook.Save("TimelineDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
