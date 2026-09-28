// Title: Reset TextBox shape text formatting to default using Aspose.Cells ClearFormatting method in C#
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates over every TextBox shape on the first worksheet, and calls ClearFormatting() to revert its text style to the default settings. | Refactor the given program to replace the manual font property assignments with a single TextBox.ClearFormatting() call while preserving the existing text content. | Create a reusable method that accepts a Worksheet object and clears the formatting of all TextBox shapes using Aspose.Cells, then show how to invoke it from a console application.
// Common Searches: Aspose.Cells C# clear formatting of all TextBox shapes in an Excel worksheet | reset shape text style to default using ClearFormatting method Aspose.Cells | how to remove rich text formatting from Excel TextBox shapes with Aspose.Cells API | C# example for iterating worksheet shapes and applying ClearFormatting in Aspose.Cells
// Tags: Aspose.Cells ClearFormatting TextBox | reset shape text formatting C# | iterate worksheet shapes Aspose.Cells | remove rich text formatting Excel Aspose.Cells | default textbox font Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, accesses the first worksheet, loops through all shapes, identifies TextBox objects, and uses the ClearFormatting() method to revert their text formatting to default values while keeping the original text unchanged, then saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only TextBox shapes (they contain text)
                if (shape is TextBox textBox)
                {
                    // Reset font formatting to typical defaults
                    textBox.Font.Color = System.Drawing.Color.Black;
                    textBox.Font.Size = 11;
                    textBox.Font.Name = "Calibri";

                    // Reassign the existing text to clear any rich‑text formatting
                    string currentText = textBox.Text;
                    textBox.Text = currentText;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
