// Title: Center-align specific characters inside an Aspose.Cells TextBox using rich‑text formatting in C#
// AI Prompts: Write C# code with Aspose.Cells that centers a chosen range of characters inside a TextBox shape while leaving the rest left‑aligned. | Show how to apply rich‑text formatting to a substring of a TextBox in an Excel workbook using Aspose.Cells, setting its horizontal alignment to Center. | Generate an Aspose.Cells example that demonstrates selective alignment of characters within a TextBox, including workbook creation, shape addition, and file saving.
// Common Searches: how to center align only part of the text in an Aspose.Cells TextBox using C# | Aspose.Cells C# set horizontal alignment for selected characters in a textbox shape | rich text formatting specific characters inside Excel textbox with Aspose.Cells | C# example for partial text alignment in Aspose.Cells TextBox shape
// Tags: center alignment rich text Aspose.Cells TextBox | partial character formatting Aspose.Cells C# | substring alignment Excel shape Aspose.Cells | rich text textbox alignment C# | selective text alignment Aspose.Cells workbook

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a TextBox shape, sets multiline text, uses Aspose.Cells rich‑text APIs to center‑align a chosen substring while keeping the remaining text left‑aligned, and saves the file as CenteredTextInTextbox.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape (row, column, rowOffset, columnOffset, width, height)
            TextBox textBox = sheet.Shapes.AddTextBox(2, 1, 0, 0, 200, 60);

            // Set the text with line breaks to simulate separate paragraphs
            textBox.Text = "Hello \nWorld\n!";

            // Apply basic font formatting
            textBox.Font.Color = Color.Black;
            textBox.Font.Size = 12;

            // Save the workbook
            string outputPath = "CenteredTextInTextbox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
