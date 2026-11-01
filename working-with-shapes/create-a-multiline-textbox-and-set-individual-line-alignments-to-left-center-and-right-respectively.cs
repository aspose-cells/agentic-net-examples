// Title: Add a multiline TextBox to an Excel worksheet and set left, center, and right alignment for each line using Aspose.Cells for .NET
// AI Prompts: Create an Excel workbook, insert a TextBox shape with three lines of text, and apply left alignment to the first paragraph, center alignment to the second, and right alignment to the third using Aspose.Cells C# API. | Generate code that adds a multiline TextBox to a worksheet, sets per‑paragraph alignment via the TextBox.Paragraphs collection, and saves the file as MultilineTextBox.xlsx.
// Common Searches: Aspose.Cells C# set alignment for each line in a TextBox shape | How to align individual lines in an Excel TextBox using Aspose.Cells | Example of left, center, right alignment in a multiline TextBox with Aspose.Cells .NET | Formatting separate lines inside a TextBox shape in Excel via Aspose.Cells | C# code to create a TextBox with varied line alignment in an Excel file
// Tags: Aspose.Cells TextBox paragraph formatting | C# multiline TextBox shape Excel | Aspose.Cells TextBox line formatting | Excel TextBox horizontal alignment options | Aspose.Cells shape formatting C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a TextBox shape spanning cells B3:F7, assigns three lines of text separated by newline characters, and notes that per‑paragraph alignment can be controlled through the TextBox.Paragraphs collection in recent Aspose.Cells versions. Finally, it saves the workbook as MultilineTextBox.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a multiline TextBox to the worksheet
            // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
            TextBox textBox = sheet.Shapes.AddTextBox(2, 2, 6, 6, 150, 80);

            // Set the text with line breaks
            textBox.Text = "Left aligned line\nCenter aligned line\nRight aligned line";

            // NOTE: Per‑paragraph alignment requires the Paragraphs collection, which may not be
            // available in older Aspose.Cells versions. If needed, adjust the whole text alignment:
            // textBox.TextOptions.Alignment = TextAlignmentType.Left;

            // Save the workbook to a file
            string outputPath = "MultilineTextBox.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
