using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

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

            // Populate worksheet with sample data (Date and Value)
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue(new DateTime(2023, 1, 1));
            cells["B2"].PutValue(100);
            cells["A3"].PutValue(new DateTime(2023, 2, 1));
            cells["B3"].PutValue(200);
            cells["A4"].PutValue(new DateTime(2023, 3, 1));
            cells["B4"].PutValue(300);

            // Add a PivotTable based on the data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable (placed at row 6, column 0)
            // Use the actual field name from the pivot to avoid mismatches
            string dateFieldName = pivot.RowFields[0].Name;
            sheet.Timelines.Add(pivot, 5, 0, dateFieldName);
            Timeline timeline = sheet.Timelines[0];

            // Apply a two‑color gradient background to the Timeline shape
            timeline.Shape.Fill.FillType = FillType.Gradient;
            GradientFill gradientFill = timeline.Shape.Fill.GradientFill;
            gradientFill.SetTwoColorGradient(
                Color.LightBlue,          // First gradient color
                Color.DarkBlue,           // Second gradient color
                GradientStyleType.Horizontal,
                1);                       // Variant

            // Rotate the Timeline shape (affects the label orientation)
            timeline.Shape.RotationAngle = 45; // degrees

            // Render the worksheet (which contains the Timeline) to a BMP image
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Bmp
            };
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);
            sheetRender.ToImage(0, "TimelineOutput.bmp");

            // Optionally save the workbook for verification
            workbook.Save("TimelineWithGradient.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}