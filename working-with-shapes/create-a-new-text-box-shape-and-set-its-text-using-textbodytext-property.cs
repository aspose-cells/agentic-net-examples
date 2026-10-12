// Title: How to add a textbox shape to an Excel worksheet and set its text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new workbook, inserts a textbox shape at a specific cell range, assigns text via the TextBody.Text property, and saves the file with Aspose.Cells. | Write a method that receives a worksheet, position coordinates, size, and a string, then adds a textbox shape at that location and sets the shape's TextBody.Text to the provided string using Aspose.Cells. | Modify an existing Excel file by programmatically adding a positioned textbox shape and updating its displayed text through the Shape.TextBody API in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add textbox shape and set its text | Set TextBody.Text for a shape in an Excel file using Aspose.Cells .NET | How to position a textbox in a worksheet with Aspose.Cells API | Insert and edit textbox content in Excel via Aspose.Cells C# example | Aspose.Cells shape TextBody property usage tutorial
// Tags: Aspose.Cells add textbox shape | Shape.TextBody.Text property C# | Excel worksheet textbox insertion Aspose.Cells | Aspose.Cells save workbook with shapes | C# Aspose.Cells shape text assignment

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// Creates a new workbook, adds a textbox shape at row 2/column 2 with defined width and height, sets its displayed text via the Shape.TextBody.Text property, and saves the workbook as TextBoxExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int rowOffset = 0;
            int columnOffset = 0;
            int width = 200;
            int height = 100;

            // AddTextBox returns a Shape object from Aspose.Cells.Drawing namespace
            Shape textBox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, rowOffset, columnOffset, width, height);

            // Set the text of the text box using TextBody.Text property
            textBox.TextBody.Text = "Hello, Aspose.Cells!";

            // Save the workbook to a file (lifecycle: save)
            workbook.Save("TextBoxExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
