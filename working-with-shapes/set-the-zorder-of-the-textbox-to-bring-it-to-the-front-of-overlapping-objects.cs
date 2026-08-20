// Title: Aspose.Cells .NET: Bring a TextBox to the Front of Overlapping Shapes (C#)
// Description: C# example that creates a workbook, adds two overlapping rectangles, inserts a textbox, and uses txtBox.ToFrontOrBack(1) to move the textbox to the top of the Z‑order before saving as ZOrderDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel shape Z-order | ToFrontOrBack | bring textbox to front | shape layering | move shape forward | overlapping shapes | Excel workbook automation
// Common Searches: Aspose.Cells bring textbox to front | C# set shape Z-order in Excel | ToFrontOrBack method example | change layering of shapes Aspose.Cells | move shape forward overlapping objects
// Developer Intent: Adjust the Z‑order of a textbox so it renders above other overlapping shapes in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Ensure annotation textboxes are visible over chart elements in generated reports. | Create diagrams where labels must overlay shapes correctly by reordering their Z‑order. | Programmatically control visual hierarchy of multiple shapes in an exported Excel file.
// AI Prompts: Show how to move a shape backward with Aspose.Cells ToFrontOrBack method. | Provide code to bring several shapes to the front in a specific order using C#. | Explain how Z-order values affect rendering of overlapping shapes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds two overlapping rectangles, inserts a textbox, and uses txtBox.ToFrontOrBack(1) to move the textbox to the top of the Z‑order before saving as ZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add two overlapping rectangles
            Shape rect1 = sheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
            rect1.Name = "Rect1";
            rect1.Text = "Back";

            Shape rect2 = sheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);
            rect2.Name = "Rect2";
            rect2.Text = "Front";

            // Add a textbox that overlaps the rectangles
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            TextBox txtBox = sheet.Shapes.AddTextBox(20, 20, 0, 0, 60, 120);
            txtBox.Name = "MyTextBox";
            txtBox.Text = "Bring to Front";

            // Bring the textbox to the front of overlapping objects
            txtBox.ToFrontOrBack(1); // Positive value moves the shape forward

            // Save the workbook
            workbook.Save("ZOrderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
