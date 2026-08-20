// Title: Rotate a TextBox Shape in Aspose.Cells for .NET and Verify Its Angle
// Description: This example creates a workbook, inserts a TextBox shape on the first worksheet, sets the shape's RotationAngle property to a chosen degree value, reads the property back to confirm the rotation, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells rotate textbox | textbox RotationAngle .NET | shape orientation Aspose.Cells | verify textbox angle | C# Excel shape rotation
// Common Searches: Aspose.Cells how to rotate a textbox | read textbox rotation angle in .NET | C# set shape rotation Aspose.Cells | check textbox orientation after rotation | example rotating Excel textbox programmatically
// Developer Intent: Apply a specific rotation to a TextBox shape and programmatically confirm that the angle was applied.
// Use Cases: Add diagonal labels to financial reports for visual emphasis. | Create angled annotations next to charts in automated Excel dashboards. | Generate custom Excel templates where text orientation varies per layout requirement.
// AI Prompts: Write C# code using Aspose.Cells to rotate a TextBox by 30 degrees and save the workbook. | Provide a function that asserts a TextBox's RotationAngle equals an expected value and throws an error if it differs. | Show how to loop through multiple TextBox shapes, assigning each a unique rotation angle with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, inserts a TextBox shape on the first worksheet, sets the shape's RotationAngle property to a chosen degree value, reads the property back to confirm the rotation, and saves the file as an .xlsx document.
    public class TextBoxRotationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, width, height
                TextBox textBox = (TextBox)worksheet.Shapes.AddTextBox(2, 2, 0, 0, 200, 100);
                textBox.Text = "Rotated TextBox";

                // Specify the rotation angle (in degrees)
                double angle = 45.0;
                textBox.RotationAngle = angle;

                // Verify the rotation by reading the property back
                double currentAngle = textBox.RotationAngle;
                Console.WriteLine("Textbox rotation angle set to: " + currentAngle);

                // Save the workbook to a file
                string outputPath = "TextBoxRotationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TextBoxRotationDemo.Run();
        }
    }
}
