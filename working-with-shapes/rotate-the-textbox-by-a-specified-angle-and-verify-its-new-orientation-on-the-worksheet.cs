// Title: How to rotate a TextBox shape by a given angle and confirm its orientation using Aspose.Cells for .NET
// AI Prompts: Add a TextBox to cell B2, assign its RotationAngle property to a variable angle (e.g., 45), and print both the expected and actual rotation values to the console. | Create a reusable C# method RotateShape(Shape shape, double angle) that sets the shape's RotationAngle and returns the applied angle. | Generate a workbook with several TextBox shapes, rotate each by a different degree, verify each shape's RotationAngle, and save the file as RotatedTextbox.xlsx. | Load an existing workbook, read the RotationAngle of a saved TextBox, and output the value to verify the orientation.
// Common Searches: Aspose.Cells C# set textbox rotation angle programmatically | how to read the rotation angle of a shape in Aspose.Cells workbook | verify that a TextBox is rotated correctly in an Excel file using Aspose.Cells | rotate multiple shapes with different angles using Aspose.Cells .NET | sample code for rotating a textbox and checking its orientation in Aspose.Cells
// Tags: Aspose.Cells rotate textbox | Aspose.Cells set shape rotation angle | Aspose.Cells read textbox rotation | C# rotate shape in worksheet | Aspose.Cells verify shape orientation | rotate multiple textboxes Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a TextBox at cell B2, sets its RotationAngle to 45 degrees, prints the expected and actual rotation values, and saves the workbook as RotatedTextbox.xlsx, demonstrating how to rotate and verify a textbox shape with Aspose.Cells for .NET.
class RotateTextboxExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define textbox position and size (zero‑based row/column, pixel offsets)
            int upperLeftRow = 1;      // Row 2
            int upperLeftColumn = 1;   // Column B
            int top = 0;               // Pixels from top of the cell
            int left = 0;              // Pixels from left of the cell
            int height = 100;          // Height in pixels
            int width = 200;           // Width in pixels

            // Add a textbox shape to the worksheet
            TextBox textbox = (TextBox)sheet.Shapes.AddTextBox(
                upperLeftRow, upperLeftColumn, top, left, height, width);
            textbox.Text = "Rotated TextBox";

            // Set rotation angle (positive = clockwise)
            double rotationAngle = 45.0;
            textbox.RotationAngle = rotationAngle;

            // Verify the rotation
            double actualAngle = textbox.RotationAngle;
            Console.WriteLine($"Expected rotation: {rotationAngle} degrees");
            Console.WriteLine($"Actual rotation:   {actualAngle} degrees");

            // Save the workbook
            workbook.Save("RotatedTextbox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
