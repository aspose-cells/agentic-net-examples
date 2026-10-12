// Title: Set the text of a specific named textbox shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Using Aspose.Cells in C#, locate the shape named "TextBox 1" on the first worksheet, assign a new string to its Text property, and save the workbook to a new file. | Write C# code that opens an existing .xlsx file, iterates through worksheet.Shapes, finds the textbox with the given Name, updates its displayed text, and persists the changes. | Show how to modify the content of a named textbox shape in an Excel file with Aspose.Cells for .NET without affecting other shapes.
// Common Searches: Aspose.Cells C# change text of a textbox shape by name in an existing workbook | How to update the displayed text of a specific textbox in an Excel file using Aspose.Cells for .NET | C# example for setting Shape.Text property of a named textbox in .xlsx | Modify textbox content in Excel using Aspose.Cells without creating a new worksheet
// Tags: Aspose.Cells set shape text property | C# update textbox shape in Excel workbook | modify named textbox Aspose.Cells | shape.Text assignment Aspose.Cells .NET | iterate worksheet shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The program loads an existing Excel workbook, searches the first worksheet for a textbox shape named "TextBox 1", assigns a provided string to its Text property, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input parameters
            string inputFilePath = "input.xlsx";      // Path to the source workbook
            string outputFilePath = "output.xlsx";    // Path where the modified workbook will be saved
            string textBoxName = "TextBox 1";         // Name of the textbox to modify
            string newText = "Your provided string value"; // Text to display in the textbox

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFilePath);

                // Assume the textbox is on the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the textbox shape by name (type check omitted to avoid ShapeType compile issue)
                foreach (Shape shape in worksheet.Shapes)
                {
                    try
                    {
                        if (shape.Name == textBoxName)
                        {
                            // Set the displayed text of the textbox
                            shape.Text = newText;
                            break; // Exit after updating the desired textbox
                        }
                    }
                    catch (Exception innerEx)
                    {
                        // Handle any shape-specific errors without stopping the loop
                        Console.WriteLine($"Error processing shape '{shape.Name}': {innerEx.Message}");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to {outputFilePath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
