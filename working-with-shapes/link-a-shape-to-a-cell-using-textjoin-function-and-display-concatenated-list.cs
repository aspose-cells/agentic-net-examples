// Title: Aspose.Cells for .NET: Link a TextBox shape to a TEXTJOIN formula (C# example)
// Description: Shows how to create a workbook, fill A1:A5 with values, use TEXTJOIN in C1 to concatenate them, add a TextBox shape, bind the shape to cell C1, style the shape, and save the file.
// Keywords: Aspose.Cells | C# example | SetLinkedCell | TextBox shape | TEXTJOIN | concatenate range | link shape to cell | Excel shape binding | dynamic label | Aspose.Cells tutorial
// Common Searches: Aspose.Cells link shape to cell C# | SetLinkedCell TEXTJOIN example | How to bind a TextBox to a formula cell in Aspose.Cells | C# Aspose.Cells shape displays concatenated list | Create dynamic label with TEXTJOIN and shape
// Developer Intent: Display a comma‑separated list of values inside a TextBox shape by linking the shape to a cell that contains a TEXTJOIN formula.
// Use Cases: Add a live label that updates automatically when the source range changes. | Build a simple dashboard element that aggregates multiple cells into a readable string. | Generate a printable report where a shape shows a consolidated list of items.
// AI Prompts: Generate C# code using Aspose.Cells to add a TextBox, set a TEXTJOIN formula, link the shape to the formula cell, and apply basic formatting. | Explain how SetLinkedCell works with formula cells and how the shape refreshes when source data is modified. | Provide an example that creates several TextBox shapes, each linked to different TEXTJOIN results, within the same workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, fill A1:A5 with values, use TEXTJOIN in C1 to concatenate them, add a TextBox shape, bind the shape to cell C1, style the shape, and save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in A1:A5
            sheet.Cells["A1"].PutValue("Apple");
            sheet.Cells["A2"].PutValue("Banana");
            sheet.Cells["A3"].PutValue("Cherry");
            sheet.Cells["A4"].PutValue("Date");
            sheet.Cells["A5"].PutValue("Elderberry");

            // Set a formula in C1 that concatenates the list using TEXTJOIN
            sheet.Cells["C1"].Formula = "=TEXTJOIN(\", \",TRUE,A1:A5)";

            // Add a TextBox shape to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, top, left, width, height
            TextBox shape = sheet.Shapes.AddTextBox(2, 0, 2, 0, 300, 50);

            // Link the shape to the cell containing the TEXTJOIN formula
            // The linked cell will display the concatenated result inside the shape
            shape.SetLinkedCell("C1", false, false);

            // Adjust the shape appearance
            shape.Fill.FillType = FillType.Solid;
            // Optional: set a solid fill color (commented out to avoid compatibility issues)
            // shape.Fill.ForeColor = Color.LightYellow;
            shape.Line.Weight = 1.0;
            // Optional: set dash style if needed (commented out to avoid compatibility issues)
            // shape.Line.DashStyle = LineDashStyle.Solid;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("LinkedShape_TextJoin.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
