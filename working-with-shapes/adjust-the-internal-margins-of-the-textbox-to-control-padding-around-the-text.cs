// Title: Adjust TextBox internal margins (padding) in Aspose.Cells for .NET
// Description: Creates a workbook, adds a TextBox shape, sets custom left, right, top, and bottom margins via TextBoxOptions, disables automatic margin calculation, and saves the file. Demonstrates precise control of textbox padding in Excel using Aspose.Cells C# API.
// Keywords: Aspose.Cells TextBox padding | custom textbox margins .NET | TextBoxOptions LeftMarginPt | disable auto margin Aspose.Cells | C# Excel shape internal margins | Aspose.Cells shape formatting
// Common Searches: how to set padding inside a textbox shape using Aspose.Cells C# | Aspose.Cells disable automatic textbox margins | change left and right margins of Excel textbox with Aspose | adjust top bottom padding of a textbox in .NET | Aspose.Cells TextBoxOptions margin example
// Developer Intent: The developer needs to apply explicit padding to a TextBox shape in an Excel worksheet and prevent Aspose.Cells from overriding those values with automatic margin calculations.
// Use Cases: Design reports where textbox content must stay a fixed distance from the shape edges for a clean layout. | Generate Excel dashboards with consistently spaced text inside shapes across multiple sheets. | Create templates that require precise control of internal margins to meet branding guidelines.
// AI Prompts: Generate C# code that sets custom left, right, top, and bottom margins for a TextBox in Aspose.Cells and disables auto‑margin. | Explain why disabling TextBody.TextAlignment.IsAutoMargin is required when defining explicit textbox padding in Aspose.Cells. | Show how to modify only the left margin of a TextBox while leaving other margins at their default values using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a TextBox shape, sets custom left, right, top, and bottom margins via TextBoxOptions, disables automatic margin calculation, and saves the file. Demonstrates precise control of textbox padding in Excel using Aspose.Cells C# API.
    class AdjustTextboxMargins
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
                TextBox textBox = sheet.Shapes.AddTextBox(1, 1, 2, 2, 200, 100);
                textBox.Text = "Text with custom padding.";

                // Set internal margins (in points) to control padding around the text
                textBox.TextBoxOptions.LeftMarginPt = 15;    // left padding
                textBox.TextBoxOptions.RightMarginPt = 15;   // right padding
                textBox.TextBoxOptions.TopMarginPt = 10;     // top padding
                textBox.TextBoxOptions.BottomMarginPt = 10;  // bottom padding

                // Disable automatic margin calculation so the custom values are used
                textBox.TextBody.TextAlignment.IsAutoMargin = false;

                // Save the workbook
                workbook.Save("TextboxMarginsDemo.xlsx");
                Console.WriteLine("Workbook saved successfully as TextboxMarginsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            AdjustTextboxMargins.Run();
        }
    }
}
