// Title: Rotate a TextBox shape in Aspose.Cells for .NET and verify its orientation
// Description: Creates a new Workbook, adds a TextBox shape to the first worksheet, sets the shape's RotationAngle property to a chosen degree value, reads the property back to confirm the change, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells | C# | textbox rotation | RotationAngle | shape orientation | Excel shape manipulation | rotate TextBox programmatically
// Common Searches: rotate textbox shape Aspose.Cells .NET | get RotationAngle of a shape in Aspose.Cells | verify textbox rotation after saving workbook | set shape orientation Aspose.Cells C#
// Developer Intent: The developer needs to apply a specific rotation to a TextBox shape and confirm that the angle is correctly stored in the workbook.
// Use Cases: Tilt a label by 45° in a generated report for visual emphasis. | Assign dynamic rotation angles to multiple textboxes based on data trends. | Check that the rotation persists after the workbook is saved and reopened.
// AI Prompts: Provide C# code that rotates a TextBox shape to 30 degrees and reads back its RotationAngle using Aspose.Cells. | Show how to loop through all TextBox shapes on a worksheet and output each shape's RotationAngle. | Explain the steps to set a shape's RotationAngle and ensure the orientation is retained in the saved Excel file with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, adds a TextBox shape to the first worksheet, sets the shape's RotationAngle property to a chosen degree value, reads the property back to confirm the change, and saves the file as an .xlsx document.
class RotateTextBoxDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(2, 2, 0, 0, 200, 100);
        textBox.Text = "Rotated TextBox";

        // Specify the rotation angle (in degrees) and apply it to the textbox
        double rotationAngle = 45; // Example angle
        textBox.RotationAngle = rotationAngle;

        // Verify the rotation by reading back the property
        Console.WriteLine("Textbox rotation angle set to: " + textBox.RotationAngle);

        // Save the workbook to a file
        workbook.Save("RotatedTextBox.xlsx");
    }
}
