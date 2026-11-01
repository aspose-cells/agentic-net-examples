// Title: Lock a specific 'Signature' drawing shape in an Excel file with Aspose.Cells for .NET to stop moving or resizing
// AI Prompts: Write C# code that locates a shape by its Name property in a worksheet and sets its IsLocked flag using Aspose.Cells. | Demonstrate how to protect a drawing object in an existing .xlsx file so it cannot be resized or repositioned with Aspose.Cells. | Show how to iterate through worksheet.Shapes, find the 'Signature' shape, enable locking, and save the workbook in C#.
// Common Searches: Aspose.Cells C# lock drawing shape from being moved | prevent users from resizing a shape in Excel using Aspose.Cells API | set IsLocked on a specific shape in a workbook with Aspose.Cells | C# find shape by name in worksheet and make it read‑only | how to protect a signature image in an Excel file programmatically
// Tags: Aspose.Cells IsLocked drawing protection | C# retrieve worksheet shape using Name property | Excel shape immutability with Aspose.Cells | prevent shape movement in .xlsx via API | signature image locking example

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load the existing workbook
Workbook workbook = new Workbook("input.xlsx");

// Access the first worksheet (adjust index if needed)
Worksheet worksheet = workbook.Worksheets[0];

// Find the shape named "Signature"
Shape signatureShape = null;
foreach (Shape shape in worksheet.Shapes)
{
    if (shape.Name == "Signature")
    {
        signatureShape = shape;
        break;
    }
}

// If the shape is found, lock it to prevent moving or resizing
if (signatureShape != null)
{
    // This property locks the shape from being moved, resized, or edited
    signatureShape.IsLocked = true;
}
else
{
    // Optionally handle the case where the shape does not exist
    System.Console.WriteLine("Shape 'Signature' not found.");
}

// Save the workbook with the changes
workbook.Save("output.xlsx");
