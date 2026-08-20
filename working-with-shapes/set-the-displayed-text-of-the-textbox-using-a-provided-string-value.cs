// Title: Set Text of a TextBox Shape in Aspose.Cells (C#)
// Description: Creates a new Workbook, adds a TextBox shape to the first worksheet, assigns a provided string to the shape's Text property, enables ResizeToFitText, and saves the file as TextboxWithText.xlsx.
// Keywords: Aspose.Cells C# textbox | set textbox text Aspose.Cells | AddTextBox shape .NET | ResizeToFitText | save workbook with textbox | textbox shape text property | Aspose.Cells example | C# Excel textbox shape
// Common Searches: Aspose.Cells set textbox text C# | how to add textbox to worksheet Aspose.Cells | ResizeToFitText Aspose.Cells example | save Excel file with textbox Aspose.Cells | C# code for textbox shape in Aspose.Cells
// Developer Intent: Add a textbox shape to a worksheet and display a supplied string inside it.
// Use Cases: Generate a report where the title or heading is placed in a dynamically sized textbox. | Insert user‑provided notes or comments into a spreadsheet via a textbox shape. | Create localized Excel files that automatically adjust textbox size to fit varying string lengths.
// AI Prompts: Write C# code that adds a TextBox shape to an Aspose.Cells worksheet, sets its Text from a variable, enables ResizeToFitText, and saves the workbook. | Show how to insert a multiline string into a textbox shape in Aspose.Cells and ensure the shape resizes to fit the content. | Explain how to modify font style and alignment of a textbox after setting its text using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, adds a TextBox shape to the first worksheet, assigns a provided string to the shape's Text property, enables ResizeToFitText, and saves the file as TextboxWithText.xlsx.
public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string content = "Hello, Aspose.Cells!";
            if (args.Length > 0)
                content = args[0];

            SetTextboxTextDemo.Run(content);
            Console.WriteLine("Workbook saved successfully as 'TextboxWithText.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class SetTextboxTextDemo
{
    // Sets the displayed text of a textbox shape using the provided string value.
    public static void Run(string textboxContent)
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet.
        // Parameters: upper left row, upper left column, top, left, width, height.
        Shape textBox = worksheet.Shapes.AddTextBox(1, 1, 100, 100, 200, 50);

        // Set the text that will be displayed inside the textbox.
        textBox.Text = textboxContent;

        // Resize the shape automatically so the whole text fits.
        textBox.TextBoxOptions.ResizeToFitText = true;

        // Save the workbook to a file.
        workbook.Save("TextboxWithText.xlsx");
    }
}
