// Title: Aspose.Cells C# – Set Text and Auto‑Resize a TextBox Shape in Excel
// Description: Demonstrates how to create a workbook, add a TextBox shape to the first worksheet, assign a custom string to the shape's Text property, enable ResizeToFitText for automatic sizing, and save the file as an .xlsx document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells set textbox text C# | Add TextBox shape Aspose.Cells | ResizeToFitText Aspose.Cells | Save workbook with shape Aspose.Cells | Excel textbox shape .NET | Aspose.Cells TextBoxOptions
// Common Searches: how to set textbox text in Aspose.Cells C# | auto resize textbox shape Aspose.Cells | add textbox to worksheet using Aspose.Cells | save Excel file with textbox Aspose.Cells | Aspose.Cells TextBoxOptions example
// Developer Intent: Insert a TextBox shape, set its displayed text, enable automatic resizing, and save the workbook with Aspose.Cells for .NET.
// Use Cases: Generate a report where section titles are placed inside auto‑sized textboxes. | Create a template that adds a labeled textbox at a specific cell location. | Update existing textbox shapes with dynamic content while keeping the shape size appropriate.
// AI Prompts: Write C# code to change the text of an existing Aspose.Cells TextBox shape and turn off ResizeToFitText. | Show how to loop through all TextBox shapes in a worksheet and populate their Text property from a dictionary. | Explain the steps to add a TextBox with custom dimensions, set its text, enable ResizeToFitText, and save the workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a TextBox shape to the first worksheet, assign a custom string to the shape's Text property, enable ResizeToFitText for automatic sizing, and save the file as an .xlsx document using Aspose.Cells for .NET.
    public class SetTextboxTextDemo
    {
        public static void Run()
        {
            try
            {
                // The text to display in the textbox
                string displayText = "Hello, Aspose.Cells!";

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape textBoxShape = worksheet.Shapes.AddTextBox(1, 1, 50, 50, 200, 100);

                // Set the displayed text of the textbox
                textBoxShape.Text = displayText;

                // Resize the shape to fit the text
                textBoxShape.TextBoxOptions.ResizeToFitText = true;

                // Define output file path
                string outputPath = "SetTextboxTextDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetTextboxTextDemo.Run();
        }
    }
}
