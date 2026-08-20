// Title: Clear TextBox Content in an Excel Worksheet using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a TextBox shape, then empties its Text property to reset the displayed content while preserving size, position, and formatting, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | clear textbox | reset textbox text | empty TextBox | Excel shape | worksheet TextBox | remove text from shape | Aspose.Cells example
// Common Searches: Aspose.Cells clear textbox text C# | How to empty a TextBox shape in Excel with Aspose.Cells | Reset TextBox content programmatically Aspose.Cells .NET | Remove text from worksheet TextBox using Aspose.Cells
// Developer Intent: Remove all text from a TextBox shape in an Excel worksheet while keeping its formatting intact.
// Use Cases: Clear placeholder text after generating a report so the next run starts with a blank box. | Reset a template TextBox before reusing the workbook for new data. | Strip user‑entered notes from a TextBox before exporting the final Excel file. | Prepare a clean workbook for downstream processing by emptying all TextBox contents.
// AI Prompts: Generate C# code that clears the Text property of every TextBox in an Aspose.Cells workbook without changing size or style. | Show how to empty a specific TextBox by index in a worksheet using Aspose.Cells for .NET. | Provide an Aspose.Cells .NET snippet that clears a TextBox, saves the workbook, and logs the full file path.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Creates a workbook, adds a TextBox shape, then empties its Text property to reset the displayed content while preserving size, position, and formatting, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox at row 1, column 1 with width 200 and height 100 (points)
                int textBoxIndex = worksheet.TextBoxes.Add(1, 1, 200, 100);

                // Retrieve the added textbox
                TextBox textBox = worksheet.TextBoxes[textBoxIndex];

                // Set initial text (optional demonstration)
                textBox.Text = "Initial content of the textbox";

                // Clear the text
                textBox.Text = string.Empty;

                // Define output file path
                string outputPath = "ClearedTextBox.xlsx";

                // Save the workbook
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
