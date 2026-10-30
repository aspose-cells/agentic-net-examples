// Title: How to locate a TextBox shape by its name and modify its text, font, and size using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find textbox shape by name in worksheet | update text and font of a named TextBox in Excel using Aspose.Cells .NET | change width and height of a specific TextBox shape with Aspose.Cells | iterate worksheet shapes to locate a TextBox named MyTextBox in C# | programmatically modify textbox properties in an Excel file with Aspose.Cells
// Tags: locate textbox by name Aspose.Cells | modify textbox text Aspose.Cells .NET | set textbox font color Aspose.Cells | adjust textbox dimensions Aspose.Cells | iterate worksheet shapes Aspose.Cells C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates the first worksheet's shapes to find a TextBox named 'MyTextBox', then updates its text, font size, font color, width, and height before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Name of the textbox shape to modify
            const string textboxName = "MyTextBox";

            // Locate the shape by name and ensure it is a TextBox
            TextBox targetTextBox = null;
            foreach (Shape shape in sheet.Shapes)
            {
                // Check name and type by using 'is' operator
                if (shape.Name == textboxName && shape is TextBox tb)
                {
                    targetTextBox = tb;
                    break;
                }
            }

            if (targetTextBox != null)
            {
                // Update textbox content and formatting
                targetTextBox.Text = "Updated text content";
                targetTextBox.Font.Size = 14;
                targetTextBox.Font.Color = Color.Blue;

                // Adjust size (optional)
                targetTextBox.Width = 200;
                targetTextBox.Height = 50;
            }
            else
            {
                Console.WriteLine($"Textbox named '{textboxName}' was not found.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
