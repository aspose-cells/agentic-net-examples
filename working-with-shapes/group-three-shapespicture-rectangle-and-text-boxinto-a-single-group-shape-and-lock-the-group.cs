// Title: Create and lock a grouped shape containing a picture, rectangle, and text box with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# group a picture, rectangle and textbox into one shape and lock it | how to lock a group of shapes in an Excel workbook using Aspose.Cells for .NET | C# add picture, rectangle, and text box then create a locked group shape with Aspose.Cells | prevent moving grouped shapes in Excel with Aspose.Cells C# IsLocked property | set freefloating placement for a grouped shape in Aspose.Cells C#
// Tags: group shapes Aspose.Cells | lock shape group Aspose.Cells | add picture shape Aspose.Cells | create rectangle shape Aspose.Cells | text box shape Aspose.Cells | freefloating placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GroupShapesExample
{
    // The example demonstrates how to add a rectangle, a text box, and an embedded picture to a worksheet, combine them into a single GroupShape, lock the group to prevent moving or resizing, assign a name and free‑floating placement, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            string imagePath = "sample.jpg";          // picture file to embed
            string outputPath = "GroupedShapes.xlsx";

            // Verify that the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Error: Image file \"{imagePath}\" not found.");
                return;
            }

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                ShapeCollection shapes = sheet.Shapes;

                // Define common dimensions (in pixels)
                int rectRow = 2, rectCol = 2, rectTop = 0, rectLeft = 0, rectHeight = 80, rectWidth = 150;
                int txtRow = 6, txtCol = 2, txtTop = 0, txtLeft = 0, txtHeight = 80, txtWidth = 150;
                int picRow = 10, picCol = 2, picHeight = 80, picWidth = 150;

                // Add a rectangle shape
                Shape rectangle = shapes.AddRectangle(rectRow, rectTop, rectCol, rectLeft, rectHeight, rectWidth);
                rectangle.Fill.FillType = FillType.Solid;
                rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightGreen;
                rectangle.Line.DashStyle = MsoLineDashStyle.Solid;
                rectangle.Line.Weight = 1;
                rectangle.Text = "Rectangle";

                // Add a text box shape
                Shape textBox = shapes.AddTextBox(txtRow, txtTop, txtCol, txtLeft, txtHeight, txtWidth);
                textBox.Text = "Sample Text Box";
                textBox.Font.Size = 12;
                textBox.Font.Color = System.Drawing.Color.Black;
                textBox.Fill.FillType = FillType.Solid;
                textBox.Fill.SolidFill.Color = System.Drawing.Color.LightYellow;

                // Add a picture shape
                Shape picture;
                using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    picture = shapes.AddPicture(picRow, picCol, fs, picHeight, picWidth);
                }
                picture.AlternativeText = "Embedded picture";

                // Group the three shapes
                Shape[] itemsToGroup = new Shape[] { rectangle, textBox, picture };
                GroupShape group = shapes.Group(itemsToGroup);

                // Lock the group so it cannot be moved or resized when the sheet is protected
                group.IsLocked = true;

                // Optionally, give the group a name and set its placement
                group.Name = "MyGroupedShape";
                group.Placement = PlacementType.FreeFloating;

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the grouped shapes:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
