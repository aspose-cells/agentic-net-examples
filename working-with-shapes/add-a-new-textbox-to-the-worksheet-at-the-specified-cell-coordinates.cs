// Title: Add a TextBox shape anchored to cell B2 with custom size and formatting using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, then insert a TextBox shape at row 1, column 1 (cell B2) with a height of 50 pt and width of 200 pt using Aspose.Cells C#. | Apply blue font color, 12‑point size, and a 1‑point line weight to the TextBox, then set its text to "Hello, Aspose.Cells!". | Save the modified workbook as an .xlsx file and handle any exceptions that may occur during shape insertion.
// Common Searches: asp.net how to anchor a textbox shape to a specific cell in an Excel file with Aspose.Cells | c# Aspose.Cells add textbox to B2 with custom dimensions and font styling | set textbox line weight and font color in Aspose.Cells workbook programmatically | exception handling when inserting shapes with Aspose.Cells for .NET
// Tags: insert textbox shape into worksheet cell Aspose.Cells C# | position textbox using row and column indices Aspose.Cells | define textbox height width offset Aspose.Cells | apply font color and size to textbox Aspose.Cells | configure textbox line weight Aspose.Cells | export workbook after adding shape Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;

// The example creates a new workbook, accesses the first worksheet, adds a TextBox shape anchored at cell B2 (row 1, column 1) with specified height, width, and pixel offsets, sets its text, font color, size, and line weight, then saves the file as Output.xlsx while handling potential exceptions.
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

            // Cell coordinates for the textbox (B2)
            int row = 1;      // zero‑based index for row 2
            int column = 1;   // zero‑based index for column B

            // Define textbox size (height and width in points)
            int height = 50;
            int width = 200;

            // Offsets from the upper‑left corner of the cell (in pixels)
            int topOffset = 0;
            int leftOffset = 0;

            // Add a textbox anchored to the specified cell
            // Parameters: upperLeftRow, upperLeftColumn, topOffset, leftOffset, height, width
            TextBox textbox = sheet.Shapes.AddTextBox(row, column, topOffset, leftOffset, height, width);

            // Set the text inside the textbox
            textbox.Text = "Hello, Aspose.Cells!";

            // Optional formatting
            textbox.Font.Color = Color.Blue;
            textbox.Font.Size = 12;
            textbox.Line.Weight = 1.0;
            // DashStyle setting removed due to unavailable enum in current Aspose.Cells version

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
