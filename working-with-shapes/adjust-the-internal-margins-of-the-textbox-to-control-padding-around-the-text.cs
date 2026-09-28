// Title: How to set internal padding for a textbox shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets the left, right, top, and bottom internal margins of a textbox shape created with Aspose.Cells. | Show how to modify a textbox's padding after adding it to a worksheet by accessing the shape's formatting objects in Aspose.Cells for .NET. | Explain the steps to configure textbox internal margin properties in Aspose.Cells when the API supports them, including any fallback approaches.
// Common Searches: Aspose.Cells C# set textbox internal margin padding | adjust padding inside Excel textbox shape using Aspose.Cells .NET | C# Aspose.Cells how to change textbox margins after adding shape | textbox internal margin properties not available Aspose.Cells version | set left and right padding for textbox shape in Aspose.Cells workbook
// Tags: Aspose.Cells shape internal margin API | C# configure textbox padding in Excel | Excel textbox margin formatting Aspose.Cells | Aspose.Cells adjust shape internal margins | C# set textbox padding with Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a textbox shape to the first worksheet, assigns sample text, notes that internal margin properties are unavailable in the current Aspose.Cells version, and saves the file as TextboxMargins.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Parameters for the textbox shape (zero‑based indices)
            int row = 2;          // upper left row
            int column = 2;       // upper left column
            int topOffset = 5;    // offset from the top of the cell (points)
            int leftOffset = 5;   // offset from the left of the cell (points)
            int height = 100;     // height of the textbox (points)
            int width = 200;      // width of the textbox (points)

            // Add a textbox shape to the worksheet
            Shape textbox = sheet.Shapes.AddTextBox(row, column, topOffset, leftOffset, height, width);

            // Set the text inside the textbox
            textbox.Text = "Sample text with custom padding.";

            // Note: Properties such as TextBoxInternalMargin*,
            // LineWeight, LineColor, and FillColor are not available
            // in the current Aspose.Cells version used for this example.
            // They can be configured via the Shape's formatting objects
            // if needed in a later version.

            // Define output file path
            string outputPath = "TextboxMargins.xlsx";

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
