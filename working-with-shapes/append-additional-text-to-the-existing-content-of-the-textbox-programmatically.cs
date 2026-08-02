// Title: Append Text to a TextBox Shape in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a TextBox shape, set initial text, and programmatically append additional characters using the FontSettingCollection.AppendText method, then save the file as AppendTextDemo.xlsx.
// Keywords: Aspose.Cells C# textbox append text | FontSettingCollection AppendText | add text to Excel shape programmatically | modify textbox content Aspose.Cells | Excel automation C# Aspose | append string to TextBox shape
// Common Searches: Aspose.Cells how to append text to a textbox | C# FontSettingCollection AppendText example | Add more text to an existing Excel textbox using Aspose | Programmatically update TextBox shape in .NET | Append string to TextBody in Aspose.Cells
// Developer Intent: The developer needs to add new text to the existing content of a TextBox shape in an Excel workbook using Aspose.Cells for .NET without overwriting the original string.
// Use Cases: Build dynamic report headers by concatenating user‑provided phrases at runtime. | Update a comment or note box with status messages as a spreadsheet is processed. | Create templates that preserve preset instructions while appending custom remarks.
// AI Prompts: Show how to append multiple lines, including line breaks, to a TextBox using FontSettingCollection in Aspose.Cells C#. | Provide code to read, prepend, replace, or clear TextBox content before appending new text. | Explain handling of Unicode or right‑to‑left languages when appending text to a TextBox shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a workbook, add a TextBox shape, set initial text, and programmatically append additional characters using the FontSettingCollection.AppendText method, then save the file as AppendTextDemo.xlsx.
class AppendTextToTextBox
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, height, width, top, left (in pixels)
        var shape = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 200);

        // Access the FontSettingCollection that holds the textbox text
        FontSettingCollection fontSettings = shape.TextBody;

        // Set initial text in the textbox
        fontSettings.Text = "Hello, ";

        // Append additional text to the existing content
        fontSettings.AppendText("World!");

        // Output the final text to the console (optional)
        Console.WriteLine("Resulting text: " + fontSettings.Text);

        // Save the workbook to a file
        workbook.Save("AppendTextDemo.xlsx");
    }
}
