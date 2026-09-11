// Title: Add a Text Box Shape to an Excel Worksheet and Set Its Text with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to insert a text box at cell C3, assign custom text via TextBody.Text, and save the workbook. | Demonstrate how to change the font size of a text box shape after setting its text with Aspose.Cells in C#. | Adapt the sample to place the text box at different row/column indices and modify its width and height.
// Common Searches: asp.net c# add textbox shape to excel file using Aspose.Cells and set its text | Aspose.Cells example for setting shape text with TextBody.Text property | how to change font size of a textbox shape in Aspose.Cells C# | position textbox shape by row and column indices Aspose.Cells | save excel workbook after inserting shapes with Aspose.Cells .NET
// Tags: add textbox shape Aspose.Cells C# | set shape text TextBody.Text Aspose.Cells | adjust textbox font size Aspose.Cells | position textbox by cell indices Aspose.Cells | save workbook with shapes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Shows how to create a new workbook, add a text box shape at a specific cell location, set its text using the TextBody.Text property, modify the font size, and save the file as an .xlsx workbook.
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

            // Define position and size for the text box (zero‑based row/column indices)
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int top = 0;    // Row offset in pixels
            int left = 0;   // Column offset in pixels
            int height = 100; // Height in points
            int width = 200;  // Width in points

            // Add a text box shape to the worksheet
            Shape textBox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, top, left, height, width);

            // Set the text of the text box
            textBox.TextBody.Text = "Hello, Aspose.Cells!";

            // Adjust font size of the text box
            try
            {
                // Shape.Font provides direct access to the font settings
                textBox.Font.Size = 12;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Font adjustment error: {ex.Message}");
            }

            // Define output file path
            string outputPath = "TextBoxExample.xlsx";

            // Save the workbook to a file
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
