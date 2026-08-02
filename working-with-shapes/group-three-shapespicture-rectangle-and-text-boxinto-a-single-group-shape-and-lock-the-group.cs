// Title: Aspose.Cells .NET: Group Picture, Rectangle & TextBox into a Locked Shape
// Description: This C# example creates a workbook, inserts a picture (or a fallback rectangle when the image file is missing), adds a rectangle and a text box, combines them into a single GroupShape, sets the group’s IsLocked flag, and writes the file as GroupedShapes.xlsx.
// Keywords: Aspose.Cells | C# | GroupShape | lock shape | picture shape from stream | rectangle shape | text box shape | shape grouping | workbook protection | .NET spreadsheet API
// Common Searches: group multiple shapes Aspose.Cells C# | lock a shape group in Excel using Aspose.Cells | add picture to worksheet from file stream Aspose.Cells | fallback shape when image not found Aspose.Cells | protect worksheet while keeping grouped graphics intact
// Developer Intent: Generate an Excel file where a picture, rectangle, and text box are combined into one immutable group.
// Use Cases: Create a branded report header that includes a logo, decorative border, and caption, all locked to prevent accidental movement. | Design a template where header elements are grouped and locked, ensuring users cannot modify them on a protected sheet. | Build a printable form where the relative positions of graphic elements stay fixed during printing or PDF export.
// AI Prompts: Write C# code with Aspose.Cells to add a picture from a file stream, a rectangle, and a text box, then group and lock them in a worksheet. | Show how to detect a missing image file when adding a picture shape and replace it with a placeholder rectangle in Aspose.Cells. | Explain how to protect an Excel worksheet while keeping a specific grouped shape locked or editable as required.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupShapesDemo
{
    // This C# example creates a workbook, inserts a picture (or a fallback rectangle when the image file is missing), adds a rectangle and a text box, combines them into a single GroupShape, sets the group’s IsLocked flag, and writes the file as GroupedShapes.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the shapes collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Path to the image file to be used as a picture shape
                string imagePath = "sampleImage.png";

                // Add a picture shape (use a stream overload and verify file existence)
                Shape picture = null;
                if (File.Exists(imagePath))
                {
                    using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    {
                        picture = shapes.AddPicture(2, 0, 2, 0, fs);
                    }
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    // Fallback: add a placeholder rectangle if the image is missing
                    picture = shapes.AddRectangle(2, 0, 2, 0, 100, 100);
                }

                // Add a rectangle shape
                Shape rectangle = shapes.AddRectangle(5, 0, 5, 0, 100, 50);

                // Add a text box shape and set its text
                Shape textBox = shapes.AddTextBox(8, 0, 8, 0, 120, 60);
                textBox.Text = "Grouped Text Box";

                // Group the three shapes into a single group shape
                List<Shape> groupItems = new List<Shape> { picture, rectangle, textBox };
                GroupShape groupShape = shapes.Group(groupItems.ToArray());

                // Lock the group so it cannot be modified when the sheet is protected
                groupShape.IsLocked = true;

                // Save the workbook with the grouped shapes
                string outputPath = "GroupedShapes.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
