using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDrawObjectRecorder
{
    public class RecordingDrawObjectHandler : DrawObjectEventHandler
    {
        public List<string> Records { get; } = new List<string>();

        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            string info = $"Type: {drawObject.Type}, Bounds: X={x}, Y={y}, W={width}, H={height}";

            if (drawObject.Cell != null)
            {
                info += $", Cell: {drawObject.Cell.Name}, Value: {drawObject.Cell.Value}";
            }

            if (drawObject.Shape != null)
            {
                info += $", Shape: {drawObject.Shape.Name}, Text: {drawObject.Shape.Text}";
            }

            Records.Add(info);
        }
    }

    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");
            sheet.Cells["A3"].Formula = "=A1 & \" \" & A2";

            var shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,   // upperLeftRow
                0,   // upperLeftColumn
                5,   // lowerRightRow
                0,   // lowerRightColumn
                120, // width
                60   // height
            );
            shape.Text = "Sample Shape";

            var pdfOptions = new PdfSaveOptions();
            var handler = new RecordingDrawObjectHandler();
            pdfOptions.DrawObjectEventHandler = handler;

            workbook.Save("RecordedOutput.pdf", pdfOptions);

            Console.WriteLine("Recorded Draw Objects:");
            foreach (var record in handler.Records)
            {
                Console.WriteLine(record);
            }
        }
    }
}