using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

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

            // Populate worksheet with sample data (dates and values)
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue(new DateTime(2023, 1, 1));
            cells["B2"].PutValue(100);
            cells["A3"].PutValue(new DateTime(2023, 2, 1));
            cells["B3"].PutValue(200);
            cells["A4"].PutValue(new DateTime(2023, 3, 1));
            cells["B4"].PutValue(300);

            // Create a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table (using the "Date" field)
            Timeline timeline = null;
            try
            {
                int timelineIdx = sheet.Timelines.Add(pivot, 10, 5, "Date");
                timeline = sheet.Timelines[timelineIdx];
            }
            catch (Exception ex)
            {
                Console.WriteLine("Timeline creation failed: " + ex.Message);
            }

            // If timeline was created, apply gradient and rotation
            if (timeline != null)
            {
                TimelineShape shape = timeline.Shape;
                shape.Fill.FillType = FillType.Gradient;
                GradientFill gradient = shape.Fill.GradientFill;
                gradient.SetTwoColorGradient(
                    Color.LightBlue,
                    Color.DarkBlue,
                    GradientStyleType.Horizontal,
                    1);
                shape.RotationAngle = 45; // degrees
            }
            else
            {
                // Fallback: create a regular rectangle shape with the same gradient
                // AddShape parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
                Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 10, 5, 10, 5, 200, 100);
                rect.Fill.FillType = FillType.Gradient;
                GradientFill gradient = rect.Fill.GradientFill;
                gradient.SetTwoColorGradient(
                    Color.LightBlue,
                    Color.DarkBlue,
                    GradientStyleType.Horizontal,
                    1);
                rect.RotationAngle = 45;
            }

            // Render the worksheet (which contains the timeline or shape) to a BMP image
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Bmp,
                OnePagePerSheet = true
            };
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "TimelineWithGradient.bmp");

            // Save the workbook (optional, to keep the timeline in the file)
            workbook.Save("TimelineWithGradient.xlsx");
        }
        catch (Exception e)
        {
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }
}