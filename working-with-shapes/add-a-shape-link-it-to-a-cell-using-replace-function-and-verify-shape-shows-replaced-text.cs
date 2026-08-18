// Title: Aspose.Cells for .NET: Add Rectangle Shape, Link to Cell, Replace Placeholder, Update Shape Text
// Description: Demonstrates creating a workbook, inserting a placeholder in A1, adding a rectangle shape, linking it to the cell, using Workbook.Replace to change the placeholder, calling Shape.UpdateSelectedValue, and saving the file so the shape shows the new text.
// Keywords: Aspose.Cells | C# | .NET | rectangle shape | linked cell | Workbook.Replace | Shape.UpdateSelectedValue | dynamic label | Excel automation | Aspose.Cells tutorial | US developers | global Excel SDK
// Common Searches: Aspose.Cells link shape to cell and refresh after replace | C# update shape text after Workbook.Replace | How to use Shape.UpdateSelectedValue in Aspose.Cells | Add rectangle shape linked to cell Aspose.Cells .NET | Replace placeholder and sync shape text Excel SDK
// Developer Intent: Add a shape, bind it to a cell, replace the cell’s placeholder text, and have the shape automatically display the updated value.
// Use Cases: Template worksheets where shapes act as live captions that change when placeholders are replaced. | Automated report generation with rectangle shapes that reflect calculated cell values instantly. | Diagram annotations that stay synchronized with data updates by linking shapes to worksheet cells.
// AI Prompts: Generate C# code with Aspose.Cells to add a rectangle shape, link it to cell A1, replace a placeholder string, and refresh the shape text. | Explain why Shape.UpdateSelectedValue is required after calling Workbook.Replace in Aspose.Cells. | Provide a step‑by‑step verification method to confirm that a shape linked to a cell shows the replaced text in the saved Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, inserting a placeholder in A1, adding a rectangle shape, linking it to the cell, using Workbook.Replace to change the placeholder, calling Shape.UpdateSelectedValue, and saving the file so the shape shows the new text.
class ShapeReplaceDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a placeholder text into a cell that will be linked to the shape
        sheet.Cells["A1"].PutValue("{{PLACEHOLDER}}");

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = sheet.Shapes.AddRectangle(1, 1, 0, 0, 150, 50);

        // Link the shape to the cell containing the placeholder
        shape.LinkedCell = "A1";

        // Optionally set some initial text for the shape (will be replaced after linking)
        shape.Text = "Initial Text";

        // Replace the placeholder in the worksheet with the desired string
        workbook.Replace("{{PLACEHOLDER}}", "Replaced Text");

        // Update the shape so that it reflects the new value of the linked cell
        shape.UpdateSelectedValue();

        // Verify that the shape now shows the replaced text
        Console.WriteLine("Shape text after replacement: " + shape.Text);

        // Save the workbook (the shape will be visible in the saved file)
        workbook.Save("ShapeReplaceDemo.xlsx");
    }
}
