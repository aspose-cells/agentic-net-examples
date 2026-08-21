// Title: Link a Shape to Cell B5 Using Hyperlink in Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape to a worksheet, set its Hyperlink.Address to "Sheet1!B5" with a ScreenTip, and save the workbook so clicking the shape jumps to cell B5.
// Keywords: Aspose.Cells shape hyperlink | C# Aspose.Cells hyperlink | link shape to cell | shape click navigation .NET | Aspose.Cells rectangle hyperlink | navigate to cell B5
// Common Searches: Aspose.Cells assign hyperlink to shape | link shape to specific cell Aspose.Cells | set ScreenTip for shape hyperlink C# | clickable shape navigate to cell B5
// Developer Intent: Create a rectangle shape and attach a hyperlink that opens cell B5 when the shape is clicked.
// Use Cases: Add a clickable button that jumps to a summary cell in a dashboard. | Create navigation links between sections of a financial report. | Provide tooltips on shapes to guide users before navigating to key data cells.
// AI Prompts: Generate C# code with Aspose.Cells to add a shape linked to cell C10 and include a custom ScreenTip. | Explain how to set a shape's Hyperlink.Address to reference another worksheet in Aspose.Cells for .NET. | Show how to programmatically create multiple shapes, each linking to different cells with unique tooltips.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a rectangle shape to a worksheet, set its Hyperlink.Address to "Sheet1!B5" with a ScreenTip, and save the workbook so clicking the shape jumps to cell B5.
    public class ShapeLinkToCellDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully: ShapeLinkToCell.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

            // Set hyperlink of the shape to cell B5 on the same sheet
            Hyperlink hyperlink = shape.Hyperlink;
            hyperlink.Address = "Sheet1!B5";          // Navigate to cell B5 when clicked
            hyperlink.ScreenTip = "Click to go to B5";

            // Save the workbook
            workbook.Save("ShapeLinkToCell.xlsx");
        }
    }
}
