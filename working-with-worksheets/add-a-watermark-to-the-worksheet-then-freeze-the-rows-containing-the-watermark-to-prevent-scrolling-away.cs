// Title: Create a diagonal semi‑transparent text watermark on an Excel worksheet and freeze the watermark rows using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add diagonal text watermark to Excel worksheet | How to freeze rows that contain a watermark in Aspose.Cells .NET | Set transparency and rotation for a text effect shape with Aspose.Cells C# | Position a free‑floating shape to cover specific rows in an Aspose.Cells workbook | Save workbook with watermark and frozen panes using Aspose.Cells for .NET
// Tags: add diagonal text watermark Aspose.Cells C# | freeze panes rows containing watermark Aspose.Cells | text effect shape transparency Aspose.Cells | free‑floating shape placement worksheet Aspose.Cells | save workbook with watermark Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWatermarkExample
{
    // The example creates a new workbook, inserts a free‑floating text effect shape as a semi‑transparent diagonal watermark covering the first three rows, adjusts its size, position, rotation, and transparency, freezes those rows, and saves the file as Watermarked.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text effect shape that will serve as the watermark
                // The method returns the created Shape object directly
                Shape watermarkShape = worksheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,   // preset effect
                    "CONFIDENTIAL",                    // watermark text
                    "Arial",                           // font name
                    48,                                // font size
                    false,                             // bold
                    false,                             // italic
                    0,                                 // left position (will be adjusted later)
                    0,                                 // top position (will be adjusted later)
                    0,                                 // width (auto)
                    0,                                 // height (auto)
                    0,                                 // shape rotation (default)
                    0);                                // shape Z‑order (default)

                // Make the shape free‑floating so it can be positioned anywhere
                watermarkShape.Placement = PlacementType.FreeFloating;

                // Position the watermark to cover the top three rows
                watermarkShape.Left = 0;
                watermarkShape.Top = 0;

                // Approximate width to span the entire worksheet (using column count)
                int approximateColumnWidth = 64; // approximate pixel width of a column
                watermarkShape.Width = (worksheet.Cells.MaxColumn + 1) * approximateColumnWidth;

                // Height to cover three rows (using row height)
                int rowHeight = (int)worksheet.Cells.Rows[0].Height;
                watermarkShape.Height = rowHeight * 3;

                // Rotate the watermark for a typical diagonal appearance
                watermarkShape.RotationAngle = -45;

                // Make the watermark semi‑transparent (use Fill property)
                watermarkShape.Fill.Transparency = 0.8; // 80% transparent

                // Freeze the rows that contain the watermark (rows 1‑3)
                // FreezePanes(row, column, freezableRows, freezableColumns)
                worksheet.FreezePanes(4, 0, 3, 0);

                // Save the workbook
                string outputPath = "Watermarked.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
