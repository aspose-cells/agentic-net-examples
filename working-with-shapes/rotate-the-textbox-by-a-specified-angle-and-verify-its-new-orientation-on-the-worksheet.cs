using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
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
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                TextBox textBox = (TextBox)worksheet.Shapes.AddTextBox(2, 2, 0, 0, 200, 100);
                textBox.Text = "Rotated TextBox";

                // Specify the rotation angle (in degrees)
                double rotationAngle = 45.0;
                textBox.RotationAngle = rotationAngle;

                // Verify the rotation by reading the property back
                Console.WriteLine("Textbox rotation angle set to: " + textBox.RotationAngle);

                // Save the workbook to a file
                workbook.Save("TextBoxRotationDemo.xlsx");
                Console.WriteLine("Workbook saved as TextBoxRotationDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TextBoxRotationDemo.Run();
        }
    }
}