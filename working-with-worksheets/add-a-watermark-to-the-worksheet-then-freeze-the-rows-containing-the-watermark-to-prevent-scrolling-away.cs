// Title: Create a semi‑transparent rotated WordArt watermark and freeze its rows in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to insert a rotated WordArt watermark in an Excel file with Aspose.Cells C# | Aspose.Cells freeze rows that contain a shape watermark in .NET | set transparency for a WordArt shape using Aspose.Cells API | freeze top three rows after adding a watermark with Aspose.Cells | add CONFIDENTIAL WordArt watermark and lock rows in Excel via C#
// Tags: add WordArt watermark Aspose.Cells C# | rotate WordArt shape Aspose.Cells | set shape transparency Aspose.Cells | freeze top rows Aspose.Cells | freeze panes after adding shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a new workbook, adds a WordArt shape with the text "CONFIDENTIAL" as a semi‑transparent, diagonally rotated watermark, sends the shape behind the cells, freezes the first three rows that contain the watermark using FreezePanes, and saves the result as WatermarkAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Optional: populate some sample data
            worksheet.Cells["A1"].PutValue("Item");
            worksheet.Cells["B1"].PutValue("Quantity");
            for (int i = 2; i <= 20; i++)
            {
                worksheet.Cells[i - 1, 0].PutValue($"Product {i - 1}");
                worksheet.Cells[i - 1, 1].PutValue(i * 5);
            }

            // -------------------------------------------------
            // Add a WordArt shape that will act as a watermark
            // -------------------------------------------------
            ShapeCollection shapes = worksheet.Shapes;

            // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            Shape watermark = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "CONFIDENTIAL",
                0,      // topRow (row index where the shape starts)
                0,      // top offset in pixels
                0,      // leftColumn (column index where the shape starts)
                0,      // left offset in pixels
                30,     // height in pixels
                500);   // width in pixels

            // Format the watermark: rotate, make semi‑transparent, hide border
            watermark.RotationAngle = -45;                     // diagonal appearance
            watermark.FillFormat.Transparency = 0.8;           // 80 % transparent
            watermark.LineFormat.IsVisible = false;           // no outline
            watermark.ZOrderPosition = 0;                     // send to back so cells are selectable

            // -------------------------------------------------
            // Freeze the rows that contain the watermark
            // -------------------------------------------------
            // The watermark is placed in rows 0‑2 (Excel rows 1‑3). Freeze the first three rows.
            // FreezePanes(row, column, freezedRows, freezedColumns)
            worksheet.FreezePanes(3, 0, 3, 0); // Freeze rows 1‑3, no frozen columns

            // Save the workbook
            string outputPath = "WatermarkAndFreeze.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
