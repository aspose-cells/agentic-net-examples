using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class TextBoxNameErrorHandling
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox and give it a name
                int tbIndex = worksheet.TextBoxes.Add(2, 2, 150, 50);
                TextBox textBox = worksheet.TextBoxes[tbIndex];
                textBox.Name = "ExistingBox";
                textBox.Text = "I exist!";

                // Attempt to retrieve a textbox by name that may not exist
                string targetName = "NonExistingBox";

                try
                {
                    // This will throw if the name is not found
                    TextBox targetBox = worksheet.TextBoxes[targetName];

                    // If no exception, the textbox was found
                    Console.WriteLine($"Found TextBox: {targetBox.Name}");
                    Console.WriteLine($"Text: {targetBox.Text}");
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.Shape)
                {
                    // Specific handling for shape-related errors (e.g., missing textbox)
                    Console.WriteLine($"Error: TextBox with name \"{targetName}\" does not exist.");
                    Console.WriteLine($"Exception Code: {ex.Code}");
                    Console.WriteLine($"Message: {ex.Message}");
                }

                // Save the workbook
                string outputPath = "TextBoxNameErrorHandling.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // General fallback for any unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TextBoxNameErrorHandling.Run();
        }
    }
}