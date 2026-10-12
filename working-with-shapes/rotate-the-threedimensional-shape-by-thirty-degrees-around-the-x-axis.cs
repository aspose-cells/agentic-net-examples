// Title: Rotate a 3D rectangle shape 30° around the X‑axis using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates an Excel workbook, adds a rectangle shape, and sets its ThreeDFormat.RotationX to 30 degrees with Aspose.Cells. | Show how to change the ThreeDFormat.RotationX property to apply any custom X‑axis rotation angle for a shape in Aspose.Cells. | Implement a reusable method that receives a Shape object and an angle, then assigns the angle to ThreeDFormat.RotationX for any 3‑D shape on a worksheet.
// Common Searches: Aspose.Cells C# set ThreeDFormat.RotationX property for a shape | example of rotating a rectangle shape on the X axis in an Excel file with Aspose.Cells | C# code to apply X‑axis 3D rotation to shapes in a worksheet using Aspose.Cells | how to enable 3D formatting and rotate shapes in Aspose.Cells .NET
// Tags: Aspose.Cells ThreeDFormat.RotationX usage | rotate 3D shape X‑axis Aspose.Cells | C# add rectangle shape with 3D formatting | Excel workbook shape rotation .NET | set shape rotation angle Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This program creates a new workbook, inserts a rectangle shape, sets its ThreeDFormat.RotationX to 30 degrees to rotate the shape around the X‑axis, and saves the file as RotatedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the sheet (this shape can be displayed in 3‑D)
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 150, 100);

            // Enable 3‑D formatting by setting rotation (no explicit Is3D property in Aspose.Cells)
            shape.ThreeDFormat.RotationX = 30;

            // Save the workbook
            workbook.Save("RotatedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
