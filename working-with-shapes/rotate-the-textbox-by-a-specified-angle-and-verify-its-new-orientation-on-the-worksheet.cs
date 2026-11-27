// Title: Rotate an Excel textbox shape to a specific angle and verify its orientation using Aspose.Cells for .NET
// Common Searches: how to set rotation angle of a textbox shape in Aspose.Cells C# | read textbox RotationAngle after saving and reloading workbook with Aspose.Cells | rotate Excel textbox programmatically with Aspose.Cells .NET example | verify shape orientation persistence in a saved Excel file using Aspose.Cells | use command line argument to specify textbox rotation angle in Aspose.Cells demo | get bounding box size of a rotated textbox shape in Aspose.Cells
// Tags: Aspose.Cells textbox rotation C# | shape RotationAngle property Aspose.Cells | Excel shape orientation persistence | programmatic shape rotation Aspose.Cells | reload workbook shape rotation verification | textbox bounding box size after rotation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace RotateTextboxDemo
{
    // The example creates a workbook, adds a textbox shape, sets its RotationAngle (default 45° or a value passed via command line), prints the angle and dimensions, saves the file, reloads it, and reads back the RotationAngle to confirm that the rotation is retained.
    class Program
    {
        static void Main(string[] args)
        {
            // Angle in degrees. Use command‑line argument if supplied, otherwise default to 45°.
            float angle = 45f;
            if (args.Length > 0 && float.TryParse(args[0], out float parsed))
            {
                angle = parsed;
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Demo";

            // Add a textbox shape.
            // Parameters: upper left row, upper left column, top offset, left offset, height, width (in pixels)
            // Here we place it at row 2, column 2 with default offsets.
            Shape textbox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);
            textbox.Text = $"Rotated {angle}°";

            // Set the rotation angle. Positive values rotate clockwise.
            textbox.RotationAngle = angle;

            // Verify the rotation by reading the property back.
            Console.WriteLine($"Textbox created with rotation angle: {textbox.RotationAngle}°");

            // Optionally, you can also verify the bounding box size after rotation.
            // The shape's Height and Width remain the same; the visual appearance is rotated.
            Console.WriteLine($"Textbox size (Height x Width): {textbox.Height} x {textbox.Width}");

            // Save the workbook to an Excel file.
            string outputPath = "RotatedTextbox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");

            // Load the workbook again to ensure the rotation persists.
            Workbook reloaded = new Workbook(outputPath);
            Shape reloadedTextbox = reloaded.Worksheets[0].Shapes[0];
            Console.WriteLine($"Reloaded textbox rotation angle: {reloadedTextbox.RotationAngle}°");

            // End of demo.
        }
    }
}
