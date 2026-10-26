// Title: C# – Catch missing TextBox shape exception and add the shape in an Aspose.Cells worksheet
// Common Searches: Aspose.Cells C# catch exception when textbox shape name is missing | how to create a textbox shape if it does not exist in a worksheet using Aspose.Cells | retrieve shape by name with fallback to creation in Aspose.Cells C# | C# Aspose.Cells add textbox shape when shape lookup fails | exception handling for missing worksheet shape in Aspose.Cells
// Tags: Aspose.Cells missing textbox shape handling | C# retrieve shape by name Aspose.Cells | create textbox shape if not found Aspose.Cells | worksheet shape exception handling C# | Aspose.Cells add textbox programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing workbook or creates a new one, attempts to get a TextBox shape named "MyTextBox" from the first worksheet, catches the ArgumentException when the shape is absent, creates a new textbox with specified dimensions and assigns the name, updates its text, and saves the workbook, all wrapped in a general exception handler for unexpected errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook (optional). If it does not exist, a new workbook will be created.
            string inputPath = "input.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook or create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a workbook with one default worksheet
                }

                // Get the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                Shape textBox = null;

                // Try to retrieve the textbox shape by name.
                try
                {
                    textBox = worksheet.Shapes["MyTextBox"];
                }
                catch (ArgumentException)
                {
                    // Shape not found – create a new textbox and assign the expected name.
                    int upperLeftRow = 2;
                    int upperLeftColumn = 2;
                    int top = 0;
                    int left = 0;
                    int width = 200;
                    int height = 100;

                    textBox = worksheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, top, left, width, height);
                    textBox.Name = "MyTextBox";
                }

                // Update the textbox content.
                textBox.Text = "Updated content";

                // Save the workbook.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
