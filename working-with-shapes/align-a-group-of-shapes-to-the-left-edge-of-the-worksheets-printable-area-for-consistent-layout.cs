using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AlignShapesToPrintableArea
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data to make the sheet printable
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["A3"].PutValue("Data 2");

            // Set a left margin (in centimeters) for the printable area
            // This margin defines where the printable area starts on the sheet
            sheet.PageSetup.LeftMargin = 2.0; // 2 cm left margin

            // Add a few shapes that we want to align to the printable area
            Shape shape1 = sheet.Shapes.AddRectangle(2, 2, 100, 50, 0, 0);
            shape1.Name = "Rect1";
            shape1.Text = "First";

            Shape shape2 = sheet.Shapes.AddOval(4, 4, 120, 60, 0, 0);
            shape2.Name = "Oval1";
            shape2.Text = "Second";

            Shape shape3 = sheet.Shapes.AddTextBox(6, 6, 150, 70, 0, 0);
            shape3.Name = "TextBox1";
            shape3.Text = "Third";

            // ------------------------------------------------------------
            // Align all shapes to the left edge of the printable area.
            // The printable area's left edge is defined by the left margin.
            // LeftToCorner property sets the horizontal offset of a shape
            // from the worksheet's left border (in pixels).
            // ------------------------------------------------------------

            // Convert left margin (cm) to pixels.
            // 1 cm = 28.3465 points, 1 point = 1/72 inch, 1 inch = 96 pixels.
            double cmToPixels = 28.3465 * 96.0 / 72.0; // ≈ 37.7953
            int leftMarginPixels = (int)Math.Round(sheet.PageSetup.LeftMargin * cmToPixels);

            // Apply the calculated offset to each shape.
            foreach (Shape shp in sheet.Shapes)
            {
                // Only adjust regular shapes (skip group shapes if any)
                if (!shp.IsGroup)
                {
                    shp.LeftToCorner = leftMarginPixels;
                }
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("AlignedShapes.xlsx");
        }
    }
}