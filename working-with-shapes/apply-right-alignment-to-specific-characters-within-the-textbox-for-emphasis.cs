// Title: Right-align text in a TextBox shape and apply bold Arial font with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to add a TextBox shape to a worksheet, set its text, align the text to the right, and apply a bold Arial font of size 12. | Generate an Aspose.Cells example that creates a workbook, inserts a textbox at a specific cell range, configures right horizontal alignment and custom font styling, then saves the file.
// Common Searches: Aspose.Cells C# how to set right horizontal alignment for textbox shape | C# Aspose.Cells set bold Arial font in textbox | example of adding and formatting a textbox in an Excel worksheet using Aspose.Cells .NET | align text to the right inside a shape with Aspose.Cells API | save workbook with formatted textbox Aspose.Cells C#
// Tags: textbox right horizontal alignment Aspose.Cells C# | textbox bold Arial font Aspose.Cells | add textbox shape to worksheet Aspose.Cells | save Excel file with formatted textbox Aspose.Cells | Aspose.Cells text alignment API

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// This C# example creates a new Workbook, adds a TextBox shape at row 2 column 1, sets its text, aligns the text horizontally to the right, applies a bold Arial font of size 12, and saves the workbook as AlignedTextBox.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset,
            // upper left column offset, width (in points), height (in points)
            TextBox textBox = sheet.Shapes.AddTextBox(2, 1, 0, 0, 200, 100);

            // Set the textbox text
            textBox.Text = "Aspose.Cells provides powerful features.";

            // Align text to the right inside the textbox
            textBox.TextHorizontalAlignment = TextAlignmentType.Right;

            // Apply bold Arial font to the textbox text
            var font = textBox.Font;
            font.Name = "Arial";
            font.Size = 12;
            font.IsBold = true;

            // Save the workbook
            workbook.Save("AlignedTextBox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
