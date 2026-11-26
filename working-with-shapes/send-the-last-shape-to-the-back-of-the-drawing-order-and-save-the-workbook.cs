// Title: Send the last added textbox to the back of the drawing order and save the workbook with Aspose.Cells for .NET
// Common Searches: how to move the last shape to the back layer in an Excel worksheet using Aspose.Cells C# | Aspose.Cells C# send textbox behind other shapes | reorder shape layering programmatically with Aspose.Cells | place a shape at the back of the drawing order using Aspose.Cells .NET | save workbook after changing shape order with Aspose.Cells
// Tags: Aspose.Cells shape layering manipulation | C# send shape to back | Aspose.Cells ToFrontOrBack method | Excel worksheet shape order control | save workbook after shape reordering Aspose.Cells | Aspose.Cells textbox layering

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeOrderDemo
{
    // The example creates a new workbook, adds a rectangle, an ellipse, and a textbox (the last shape) to the first worksheet, retrieves the Shapes collection, identifies the last shape, uses the ToFrontOrBack(-1) method to send it to the back of the drawing order, and then saves the workbook as 'LastShapeToBack.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape
                Shape rect = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 150);
                rect.Name = "Rectangle1";

                // Add an ellipse shape using AutoShapeType.Oval
                Shape ellipse = worksheet.Shapes.AddAutoShape(AutoShapeType.Oval, 5, 0, 5, 0, 150, 150);
                ellipse.Name = "Ellipse1";

                // Add a textbox (the last shape we add)
                int textBoxIndex = worksheet.TextBoxes.Add(8, 1, 200, 80);
                TextBox textBox = worksheet.TextBoxes[textBoxIndex];
                textBox.Name = "TextBox1";
                textBox.Text = "Last shape - will be sent to back";

                // Retrieve the shapes collection (TextBox is also represented as a Shape)
                ShapeCollection shapes = worksheet.Shapes;
                if (shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                    return;
                }

                // Get the last shape (the textbox)
                Shape lastShape = shapes[shapes.Count - 1];
                Console.WriteLine($"Last shape before reordering: Name = {lastShape.Name}, Type = {lastShape.GetType().Name}");

                // Send the last shape to the back of the drawing order
                lastShape.ToFrontOrBack(-1);
                Console.WriteLine("Last shape has been sent to the back of the drawing order.");

                // Save the workbook
                string outputPath = "LastShapeToBack.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
