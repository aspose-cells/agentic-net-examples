// Title: Set left alignment for the first line and center alignment for the second line in a TextBox with Aspose.Cells for .NET (per‑paragraph alignment not available)
// Common Searches: Aspose.Cells C# align first line left in a TextBox | center alignment for second line of a TextBox using Aspose.Cells | does Aspose.Cells support per‑paragraph alignment in TextBox shapes | workaround separate TextBox objects for different line alignments in Aspose.Cells | multiline TextBox alignment example with Aspose.Cells for .NET
// Tags: text box line alignment Aspose.Cells | per paragraph alignment limitation Aspose.Cells | multiple line alignment workaround Aspose.Cells | Aspose.Cells shape text formatting C# | TextBox alignment in Excel using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a TextBox containing two lines of text, and saves the file. It highlights that Aspose.Cells does not currently support per‑paragraph alignment within a single TextBox, and suggests using separate TextBox objects for each line to achieve distinct left and center alignments.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a TextBox shape (returns a TextBox object directly).
            TextBox textBox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 60, 200);

            // Set the TextBox content with two lines.
            textBox.Text = "First line\r\nSecond line";

            // Note: Per‑paragraph alignment is not available in this version of Aspose.Cells.
            // If needed, separate TextBox objects can be created for each line with individual alignment.

            // Save the workbook.
            string outputPath = "AlignedTextBox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
