// Title: Rotate a TextBox shape to a specific angle and verify its orientation with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# rotate textbox shape 45 degrees | How to set rotation angle for a shape in an Excel workbook using Aspose.Cells | C# verify textbox rotation angle programmatically with Aspose.Cells | Read shape RotationAngle property after rotating textbox in Aspose.Cells | Save rotated textbox in Excel file using Aspose.Cells for .NET
// Tags: textbox shape rotation Aspose.Cells .NET | set shape rotation angle Excel Aspose.Cells | verify textbox orientation C# | add textbox to worksheet Aspose.Cells | read shape rotation property Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a TextBox shape to the first worksheet, sets its RotationAngle to a desired degree, reads the property back to confirm the rotation, and saves the workbook as RotatedTextbox.xlsx.
class RotateTextboxExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet.
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            TextBox textbox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);
            textbox.Text = "Rotated TextBox";

            // Specify the rotation angle in degrees (e.g., 45 degrees)
            double rotationAngle = 45.0;
            textbox.RotationAngle = rotationAngle;

            // Verify the new orientation by checking the RotationAngle property
            double actualAngle = textbox.RotationAngle;
            if (Math.Abs(actualAngle - rotationAngle) < 0.001)
            {
                Console.WriteLine($"Textbox rotation verified: {actualAngle} degrees.");
            }
            else
            {
                Console.WriteLine($"Rotation verification failed. Expected: {rotationAngle}, Actual: {actualAngle}");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("RotatedTextbox.xlsx");
            Console.WriteLine("Workbook saved as RotatedTextbox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
