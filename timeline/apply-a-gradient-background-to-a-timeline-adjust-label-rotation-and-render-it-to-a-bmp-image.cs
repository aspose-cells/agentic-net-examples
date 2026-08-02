// Title: C# – Apply Gradient Fill and Rotation to a Pivot Timeline and Export as BMP with Aspose.Cells
// Description: Creates a workbook, adds date/value data, builds a pivot table, links a timeline to the Date field, applies a horizontal two‑color gradient, rotates the timeline shape, and renders the sheet to a BMP image using Aspose.Cells for .NET.
// Keywords: Aspose.Cells timeline gradient | C# timeline rotation | export timeline to BMP | pivot table timeline Aspose | gradient fill shape Aspose.Cells | timeline rendering Aspose.Cells
// Common Searches: how to set gradient background on a timeline Aspose.Cells | rotate timeline shape C# Aspose.Cells | export worksheet with timeline to BMP | create pivot table and timeline programmatically Aspose.Cells | gradient fill options for timeline shape
// Developer Intent: Generate a pivot‑linked timeline, style it with a two‑color gradient, rotate the shape, and save the worksheet as a BMP image.
// Use Cases: Design a dashboard where the timeline background reflects a color scale and is angled for clearer date labels, then export to BMP for documentation. | Automate report generation that includes a styled timeline and provides a BMP snapshot for legacy systems that require bitmap images. | Create printable slide assets by rendering a pivot‑driven timeline with custom gradient and rotation to BMP format.
// AI Prompts: Show how to change the gradient colors and direction of a timeline shape in Aspose.Cells C#. | Provide code to export a worksheet containing a timeline to PNG while preserving gradient and rotation. | Explain how to dynamically adjust the rotation angle of a timeline based on data range in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, adds date/value data, builds a pivot table, links a timeline to the Date field, applies a horizontal two‑color gradient, rotates the timeline shape, and renders the sheet to a BMP image using Aspose.Cells for .NET.
class TimelineGradientExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(300);

            // Add a pivot table based on the data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table
            int timelineIndex = sheet.Timelines.Add(pivot, 5, 0, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Apply a two‑color gradient background to the timeline shape
            timeline.Shape.Fill.FillType = FillType.Gradient;
            GradientFill gradientFill = timeline.Shape.Fill.GradientFill;
            gradientFill.SetTwoColorGradient(
                Color.LightBlue,          // First gradient color
                Color.DarkBlue,           // Second gradient color
                GradientStyleType.Horizontal,
                1);                       // Variant

            // Rotate the timeline (affects its label orientation)
            timeline.Shape.RotationAngle = 45;

            // Render the worksheet (including the timeline) to a BMP image
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Bmp,
                OnePagePerSheet = true
            };
            SheetRender sheetRender = new SheetRender(sheet, renderOptions);
            sheetRender.ToImage(0, "TimelineOutput.bmp");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
