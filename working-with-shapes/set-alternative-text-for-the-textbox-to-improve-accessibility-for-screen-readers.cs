// Title: Add alternative text to a textbox shape in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to set alt text for a textbox shape in Excel | adding accessibility description to Excel textbox using Aspose.Cells .NET | set AlternativeText property on shape with Aspose.Cells in C# example | Aspose.Cells tutorial for screen reader support on Excel shapes | write alternative text for Excel textbox programmatically with Aspose.Cells
// Tags: Aspose.Cells set textbox AlternativeText .NET | C# Excel shape accessibility with Aspose.Cells | Aspose.Cells shape AlternativeText property example | Excel textbox alt text using Aspose.Cells | Aspose.Cells accessibility features for shapes

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a textbox shape, assign visible text, set the AlternativeText property for screen‑reader accessibility, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Add a textbox shape (row, column, offsetX, offsetY, height, width)
        var textbox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);

        // Set visible text inside the textbox
        textbox.Text = "Sample Text";

        // Set alternative text for screen readers (accessibility)
        textbox.AlternativeText = "This textbox contains sample text for demonstration purposes.";

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
