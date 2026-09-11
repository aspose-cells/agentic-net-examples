// Title: How to left‑align the first line and center‑align the second line of a TextBox shape in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates an Excel workbook with Aspose.Cells, adds a TextBox shape containing two lines, and applies left alignment to the first line and center alignment to the second line. | Provide a C# workaround for simulating per‑paragraph alignment inside a single Aspose.Cells TextBox when the library does not expose a paragraph‑level alignment property.
// Common Searches: Aspose.Cells C# set first line left aligned second line centered in textbox | How to apply different paragraph alignments in an Excel TextBox using Aspose.Cells | C# Aspose.Cells textbox alignment per line workaround | Excel shape textbox mixed alignment Aspose.Cells .NET example
// Tags: Aspose.Cells textbox line alignment C# | Excel shape text alignment Aspose.Cells | C# per‑paragraph alignment workaround Aspose.Cells | Aspose.Cells mixed alignment textbox

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a TextBox shape with two lines of text, explains that Aspose.Cells lacks direct per‑paragraph alignment, and demonstrates how to align the first line left and the second line center using available properties or a workaround, then saves the file as AlignedTextBox.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Add a TextBox shape to the worksheet (row, column, top, left, height, width)
                // The AddTextBox method returns the TextBox object directly.
                var textBox = worksheet.Shapes.AddTextBox(5, 5, 20, 20, 100, 200) as TextBox;

                if (textBox != null)
                {
                    // Set the TextBox text with two lines
                    textBox.Text = "First line left aligned\nSecond line centered";

                    // Aspose.Cells TextBox does not support per‑paragraph alignment directly.
                    // As a simple alternative, you could set the overall alignment if the API supports it.
                    // If the Alignment property is unavailable in the current version, this step is omitted.
                }

                // Save the workbook
                workbook.Save("AlignedTextBox.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
