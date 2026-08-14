// Title: C# Example: Group a Rectangle, TextBox, and Picture into a Locked GroupShape with Aspose.Cells for .NET
// Description: Demonstrates how to create a new workbook, add a rectangle, a text box, and an optional picture, combine them into a GroupShape using ShapeCollection.Group, set the group's IsLocked flag, and save the file as GroupedShapes.xlsx. Includes error handling for missing images and notes on worksheet protection.
// Keywords: Aspose.Cells | C# | .NET | GroupShape | lock group shape | shape collection | add rectangle Aspose.Cells | add textbox Aspose.Cells | add picture shape | Excel workbook example | worksheet protection | GitHub source code | Aspose.Cells tutorial
// Common Searches: Aspose.Cells group multiple shapes C# | lock GroupShape in Excel using Aspose.Cells | add picture to worksheet and group with rectangle | ShapeCollection.Group example .NET | prevent moving grouped shapes in protected sheet
// Developer Intent: Create a locked GroupShape that contains a rectangle, a text box, and a picture on an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Bundle decorative elements (image, shape, label) into a single object for easy repositioning. | Secure a composite logo or badge so users cannot edit or move its parts when the sheet is protected. | Simplify layout adjustments in reports by treating multiple shapes as one entity.
// AI Prompts: Write C# code that adds a circle, a caption, and an image to a worksheet, groups them with Aspose.Cells, and locks the group. | Explain how to protect a worksheet so that a locked GroupShape cannot be moved or edited. | Show how to add several picture shapes to a GroupShape dynamically and set the group's IsLocked property in Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new workbook, add a rectangle, a text box, and an optional picture, combine them into a GroupShape using ShapeCollection.Group, set the group's IsLocked flag, and save the file as GroupedShapes.xlsx. Includes error handling for missing images and notes on worksheet protection.
class GroupShapesDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add a rectangle shape
            Shape rectangle = shapes.AddRectangle(2, 0, 2, 0, 100, 50);

            // Add a text box shape and set its text
            Shape textBox = shapes.AddTextBox(5, 0, 5, 0, 120, 60);
            textBox.Text = "Sample Text";

            // Prepare a list of shapes to be grouped
            List<Shape> shapesToGroup = new List<Shape> { rectangle, textBox };

            // Add a picture shape if the file exists
            string picturePath = "sample.png";
            if (File.Exists(picturePath))
            {
                try
                {
                    using (FileStream pictureStream = new FileStream(picturePath, FileMode.Open, FileAccess.Read))
                    {
                        // AddPicture overload expects a Stream
                        Shape picture = shapes.AddPicture(0, 0, 0, 0, pictureStream);
                        picture.Width = 80;
                        picture.Height = 80;
                        shapesToGroup.Add(picture);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to add picture shape. {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Warning: Image file '{picturePath}' not found. Skipping picture shape.");
            }

            // Group the collected shapes if more than one shape is present
            if (shapesToGroup.Count > 1)
            {
                GroupShape group = shapes.Group(shapesToGroup.ToArray());
                group.IsLocked = true; // Lock the group when the sheet is protected
            }

            // Save the workbook
            workbook.Save("GroupedShapes.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
