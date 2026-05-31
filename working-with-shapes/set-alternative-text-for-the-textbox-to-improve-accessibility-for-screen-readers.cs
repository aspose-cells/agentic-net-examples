using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class TextBoxAlternativeTextDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                Shape textBox = worksheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);

                // Set the visible text inside the textbox
                textBox.Text = "Sample content";

                // Set alternative text to improve accessibility for screen readers
                textBox.AlternativeText = "This textbox contains sample content for screen readers";

                // Save the workbook
                string outputPath = "TextBoxAlternativeTextDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            TextBoxAlternativeTextDemo.Run();
        }
    }
}