using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class SetTextboxTextDemo
    {
        public static void Run()
        {
            try
            {
                // Text to display in the textbox
                string displayText = "Hello, Aspose.Cells!";

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape textBox = worksheet.Shapes.AddTextBox(1, 1, 50, 50, 200, 100);

                // Set the displayed text of the textbox
                textBox.Text = displayText;

                // Save the workbook
                string outputPath = "SetTextboxTextDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}