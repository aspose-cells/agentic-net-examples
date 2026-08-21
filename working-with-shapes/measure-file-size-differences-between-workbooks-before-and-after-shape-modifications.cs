// Title: C# Example: Measure Excel Workbook Size Before and After Shape Modification with Aspose.Cells
// Description: This sample creates a workbook, adds a text‑box shape, saves the file to a memory stream to capture its byte size, updates the shape’s text and calls FitToTextSize, saves again, and outputs the size before, after, and the difference, demonstrating how shape changes affect .xlsx file size.
// Keywords: Aspose.Cells | .NET | C# | Excel workbook size | shape modification | FitToTextSize | memory stream | file size comparison | text box shape | size difference
// Common Searches: How to get Excel file size using Aspose.Cells C# | Does changing shape text affect .xlsx size | Measure workbook size before and after shape resize | Aspose.Cells FitToTextSize file size impact | C# code to compare workbook byte length
// Developer Intent: Determine the byte‑size change of an .xlsx file caused by editing a shape’s content or dimensions with Aspose.Cells.
// Use Cases: Validate that shape edits keep the workbook within cloud‑storage size limits. | Benchmark the impact of different text lengths in shapes on final file size. | Log size differences automatically when updating shapes across multiple worksheets. | Optimize shape content to reduce overall workbook size in bulk processing.
// AI Prompts: Write a C# method that returns the workbook size before and after applying FitToTextSize to a shape and prints the difference. | Explain which OpenXML parts grow when a shape’s text is increased and why FitToTextSize can enlarge the .xlsx file. | Provide a loop that iterates over all worksheets, modifies each shape, captures size changes, and writes the results to a CSV file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This sample creates a workbook, adds a text‑box shape, saves the file to a memory stream to capture its byte size, updates the shape’s text and calls FitToTextSize, saves again, and outputs the size before, after, and the difference, demonstrating how shape changes affect .xlsx file size.
class ShapeSizeDifferenceDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape textBox = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 200, 50);
        textBox.Text = "Initial text";

        // Save the workbook to a memory stream to capture the size before modification
        MemoryStream beforeStream = workbook.SaveToStream(); // lifecycle rule
        long sizeBefore = beforeStream.Length;

        // Modify the shape: add more text and adjust size to fit the text
        textBox.Text += " Adding more content to increase shape size.";
        textBox.FitToTextSize();

        // Save the workbook again to capture the size after modification
        MemoryStream afterStream = workbook.SaveToStream(); // lifecycle rule
        long sizeAfter = afterStream.Length;

        // Output the size information
        Console.WriteLine($"Size before modification: {sizeBefore} bytes");
        Console.WriteLine($"Size after modification: {sizeAfter} bytes");
        Console.WriteLine($"Difference: {sizeAfter - sizeBefore} bytes");

        // Save the final workbook to a file (optional)
        workbook.Save("ShapeModificationResult.xlsx");
    }
}
