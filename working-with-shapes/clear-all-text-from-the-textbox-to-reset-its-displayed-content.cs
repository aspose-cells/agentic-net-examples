using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for TextBox class

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Optional template file – load only if it exists
                string templatePath = "template.xlsx";
                Workbook workbook = File.Exists(templatePath)
                    ? new Workbook(templatePath)
                    : new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox (row, column, height, width)
                int textboxIndex = worksheet.TextBoxes.Add(1, 1, 100, 200);
                TextBox textBox = worksheet.TextBoxes[textboxIndex];

                // Set initial text (demonstration)
                textBox.Text = "Sample text that will be cleared.";

                // Clear the text content of the textbox
                textBox.Text = string.Empty;

                // Save the workbook
                string outputPath = "ClearTextBoxDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}