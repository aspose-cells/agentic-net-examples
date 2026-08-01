// Title: C# – Enable Text Wrapping in an Aspose.Cells TextBox Shape (.NET)
// Description: Demonstrates how to add a TextBox shape to a worksheet, assign a long string, activate the WrapTextInShape option, optionally auto‑fit the row and column, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells TextBox wrap text | WrapTextInShape .NET | C# Aspose.Cells shape text wrapping | AutoFitRow AutoFitColumn Aspose.Cells | Excel textbox long text Aspose
// Common Searches: Aspose.Cells enable text wrapping in textbox | WrapTextInShape property example C# | AutoFit row after wrapping text Aspose.Cells | How to wrap long text inside a shape using Aspose.Cells | C# Aspose.Cells textbox wrap and autofit
// Developer Intent: Turn on automatic line wrapping for long text inside a TextBox shape in an Excel workbook with Aspose.Cells for .NET.
// Use Cases: Generate reports with paragraph‑style notes that stay within fixed‑size textboxes. | Create templates where long headings automatically wrap to preserve layout. | Add multi‑line comment boxes to worksheets without manual cell resizing.
// AI Prompts: Provide C# code that creates a TextBox shape, sets a long string, enables WrapTextInShape, and auto‑fits the related row and column with Aspose.Cells. | Explain the interaction between WrapTextInShape, AutoFitRow, and AutoFitColumn in Aspose.Cells for .NET. | Step‑by‑step guide to enable text wrapping in a shape and ensure the worksheet displays the wrapped content correctly.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a TextBox shape to a worksheet, assign a long string, activate the WrapTextInShape option, optionally auto‑fit the row and column, and save the workbook using Aspose.Cells for .NET.
class EnableTextboxWrapping
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a TextBox shape (parameters: upper left row, upper left column, top, left, width, height)
        TextBox textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 200, 300, 100);

        // Set a long text that needs wrapping
        textBox.Text = "This is a very long sentence that should automatically wrap onto multiple lines within the textbox shape when wrapping is enabled.";

        // Enable text wrapping inside the shape
        textBox.TextBoxOptions.WrapTextInShape = true;

        // Optional: adjust row height/column width to display wrapped text
        worksheet.AutoFitRow(0);
        worksheet.AutoFitColumn(0);

        // Save the workbook
        workbook.Save("TextboxWrapDemo.xlsx");
    }
}
