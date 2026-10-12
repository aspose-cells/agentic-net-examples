// Title: How to replace <DATE> tokens in all TextBox shapes of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, iterates through every worksheet and each Shape, detects TextBox shapes containing the <DATE> token, and substitutes it with DateTime.Now formatted as yyyy-MM-dd. | Create a C# routine using Aspose.Cells to batch‑process an Excel file, updating the Text property of all TextBox shapes that include <DATE> with the current system date in yyyy‑MM‑dd format.
// Common Searches: aspocells c# substitute <DATE> token in textbox shapes across all worksheets | batch update text in Excel TextBox shapes using Aspose.Cells .NET | how to insert current date into shape text with Aspose.Cells C# | iterate through workbook shapes and modify placeholder text Aspose.Cells example | replace token in Excel shape text with system date Aspose.Cells tutorial
// Tags: Aspose.Cells batch shape text replacement | Excel TextBox shape date insertion C# | Iterate workbook shapes Aspose.Cells .NET | Update shape Text property with system date | Format yyyy-MM-dd in Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, loops through each worksheet and its shapes, finds TextBox shapes containing the <DATE> token, replaces the token with the current system date formatted yyyy-MM-dd, and saves the updated file.
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare the current date string in the required format
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process shapes that contain text (e.g., TextBox)
                    string text = shape.Text;
                    if (!string.IsNullOrEmpty(text) && text.Contains("<DATE>"))
                    {
                        string newText = text.Replace("<DATE>", currentDate);
                        shape.Text = newText;
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
