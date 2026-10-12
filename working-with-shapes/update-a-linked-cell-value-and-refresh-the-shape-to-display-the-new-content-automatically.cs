// Title: Update a linked cell and automatically refresh linked TextBox shapes in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to change the value of a specific cell and then refresh all TextBox shapes that reference that cell via the LinkedCell property. | Create a method that iterates through Worksheet.Shapes, identifies TextBox objects with a non‑empty LinkedCell, and updates their Text property to match the current cell value, including robust error handling. | Show how to save the modified workbook to a new file while ensuring the output directory exists, using Aspose.Cells in a .NET console application.
// Common Searches: Aspose.Cells C# update cell and refresh linked textbox shapes | How to sync Excel TextBox linked to a cell using Aspose.Cells .NET | Refresh all shapes after changing a cell value with Aspose.Cells | C# example for updating linked cell and auto‑updating textbox in Excel | Aspose.Cells iterate worksheet shapes and update linked cell content
// Tags: update cell value Aspose.Cells C# | refresh linked textbox Aspose.Cells | iterate worksheet shapes Aspose.Cells | handle linked cell errors Aspose.Cells | save workbook to new file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, changes the value of cell A1, then loops through all shapes on the first worksheet. For each TextBox with a LinkedCell reference, it reads the current cell value and assigns it to the TextBox's Text property, handling any exceptions. It also ensures the output directory exists before saving the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index or name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Update the linked cell value (example: cell A1)
                worksheet.Cells["A1"].PutValue("New Value");

                // Refresh all TextBox shapes that are linked to cells so they display the updated content
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape is TextBox textBox && !string.IsNullOrEmpty(textBox.LinkedCell))
                    {
                        try
                        {
                            // Retrieve the current value of the linked cell
                            Cell linkedCell = worksheet.Cells[textBox.LinkedCell];
                            // Update the TextBox text to match the cell value
                            textBox.Text = linkedCell.StringValue;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to refresh shape linked to {textBox.LinkedCell}: {ex.Message}");
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
