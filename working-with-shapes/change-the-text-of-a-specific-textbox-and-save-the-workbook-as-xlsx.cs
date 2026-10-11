// Title: Update the text of a specific TextBox shape in an Excel workbook and save it as XLSX with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that opens an existing XLSX file, finds a TextBox shape named 'MyTextBox' on the first worksheet, sets its Text property to a supplied string, and saves the workbook to a new XLSX file. | Refactor the sample to accept the input file path, output file path, and new TextBox content as command‑line arguments while preserving the shape‑text update logic.
// Common Searches: aspnet change textbox shape text in existing Excel file using Aspose.Cells | c# locate TextBox by name in workbook and modify its content | how to save Excel workbook after editing shape with Aspose.Cells | update specific TextBox in Excel and export as XLSX programmatically
// Tags: Aspose.Cells edit TextBox shape text | C# find shape by name in Excel workbook | Aspose.Cells save workbook as XLSX after modification | update Excel shape content programmatically | Aspose.Cells TextBox manipulation example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an existing XLSX file, searches for a TextBox shape named "MyTextBox" (or falls back to the first shape), updates its Text property, and saves the modified workbook to a new XLSX file.
class Program
{
    static void Main()
    {
        // Paths to the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate a TextBox shape (by name if possible, otherwise first shape)
            Shape? textBox = null;

            // Try to find by specific name first
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.Name == "MyTextBox")
                {
                    textBox = shape;
                    break;
                }
            }

            // If not found by name, fall back to the first shape on the sheet
            if (textBox == null && sheet.Shapes.Count > 0)
            {
                textBox = sheet.Shapes[0];
            }

            // Update the TextBox text if a shape was found
            if (textBox != null)
            {
                textBox.Text = "New text for the TextBox";
                Console.WriteLine("TextBox text updated.");
            }
            else
            {
                Console.WriteLine("No shape found in the worksheet.");
            }

            // Save the workbook as XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
