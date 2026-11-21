// Title: Append additional text to a TextBox shape in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add text to an existing textbox shape in Excel | append string to textbox shape using Aspose.Cells .NET API | modify textbox content in an Excel worksheet programmatically with Aspose.Cells | how to update the Text property of a shape in Aspose.Cells C# | concatenate text to a textbox in an existing .xlsx using Aspose.Cells
// Tags: append text to textbox shape Aspose.Cells .NET | update shape Text property Excel C# | modify textbox content Aspose.Cells workbook | Excel shape manipulation Aspose.Cells | concatenate textbox string Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing .xlsx file, finds the first shape (treated as a TextBox), reads its current Text, appends a predefined string, writes the combined text back to the shape, and saves the workbook as a new file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the first shape on the worksheet (assumed to be a TextBox)
            Shape textBox = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                // Directly take the first shape; Aspose.Cells treats TextBox as a Shape
                textBox = shape;
                break;
            }

            if (textBox != null)
            {
                try
                {
                    // Retrieve current text (null‑coalesced to empty string)
                    string currentText = textBox.Text ?? string.Empty;

                    // Define the text to append
                    const string additionalText = " Appended text.";

                    // Set the new combined text back to the textbox
                    textBox.Text = currentText + additionalText;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error updating textbox text: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No shape found on the first worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
