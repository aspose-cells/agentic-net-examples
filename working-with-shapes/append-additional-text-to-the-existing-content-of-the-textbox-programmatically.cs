// Title: How to programmatically append custom text to all TextBox shapes in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file, iterates through every TextBox shape on the first worksheet, and appends a specified string to its current Text property using Aspose.Cells. | Provide a C# snippet that prepends a custom prefix to a TextBox identified by its name while preserving other shape attributes with Aspose.Cells. | Show how to update the Text of all TextBox shapes in a worksheet without altering their formatting, then save the modified workbook with Aspose.Cells.
// Common Searches: c# aspose.cells add suffix to textbox shape in existing excel file | how to modify text of textbox objects in a worksheet using Aspose.Cells .NET | iterate over shapes collection and change textbox content aspose.cells | append custom string to every textbox in an Excel workbook programmatically | asp.net read and update textbox shape text in xlsx with Aspose.Cells
// Tags: append text to TextBox shape Aspose.Cells | iterate worksheet shapes C# | modify textbox content Excel Aspose.Cells | preserve shape formatting while updating text .NET | save updated workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, loops through all shapes on the first worksheet, identifies TextBox shapes, concatenates a custom suffix to each TextBox's existing text, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only TextBox shapes
                    if (shape is TextBox textBox)
                    {
                        // Append additional text to the existing content
                        string existingText = textBox.Text;
                        string additionalText = " – Appended text";
                        textBox.Text = existingText + additionalText;
                    }
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
}
