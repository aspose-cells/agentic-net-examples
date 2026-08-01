// Title: Clear TextBox content in an Aspose.Cells for .NET workbook (C#)
// Description: Demonstrates how to add a TextBox shape to a worksheet, set initial text, and then clear the text by assigning an empty string to the TextBox.Text property before saving the Excel file.
// Keywords: Aspose.Cells | C# | Clear TextBox | TextBox.Text empty | Excel shape text removal | Worksheet TextBox | Aspose.Cells .NET API | reset textbox content | remove shape text | clear textbox programmatically
// Common Searches: clear textbox text Aspose.Cells C# | set TextBox.Text to empty Aspose.Cells | remove text from Excel textbox using Aspose | reset shape text in Aspose.Cells | Aspose.Cells clear shape content
// Developer Intent: Remove existing text from a TextBox shape while keeping its size, position, and formatting unchanged.
// Use Cases: Prepare a template workbook by deleting placeholder text in TextBox controls before distribution. | Refresh a data‑entry sheet by clearing previous entries from TextBox shapes for a new collection cycle. | Automate cleanup of Excel reports, removing all TextBox content while preserving visual layout.
// AI Prompts: Show C# code to clear the text of a single TextBox shape using Aspose.Cells. | Provide a snippet that iterates over all TextBox objects in a worksheet and empties their Text property. | Explain how to reset TextBox content without affecting its formatting, size, or position in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a TextBox shape to a worksheet, set initial text, and then clear the text by assigning an empty string to the TextBox.Text property before saving the Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox at row 1, column 1 with width 200 and height 100 (points)
                int textboxIndex = worksheet.TextBoxes.Add(1, 1, 200, 100);
                TextBox textBox = worksheet.TextBoxes[textboxIndex];

                // Set initial text
                textBox.Text = "Sample text that will be cleared.";

                // Clear the text content
                textBox.Text = string.Empty;

                // Define output file path
                string outputPath = "ClearTextBoxContent.xlsx";

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
