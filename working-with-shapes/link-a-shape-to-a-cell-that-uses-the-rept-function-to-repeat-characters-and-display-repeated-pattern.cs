// Title: Link a Rectangle Shape to a REPT Formula Cell with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, set a REPT("*",10) formula in cell A1, add a rectangle shape, link the shape to the formula cell using SetLinkedCell, refresh the displayed value, and save the file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# shape linked cell | REPT function | SetLinkedCell | rectangle shape | dynamic shape text | Excel automation .NET | update shape value
// Common Searches: Aspose.Cells link shape to cell REPT | C# set linked cell for rectangle shape | how to display REPT formula result in a shape | update shape text from formula Aspose.Cells | bind shape to cell value .NET
// Developer Intent: Create a shape that automatically shows the text produced by a REPT formula and stays synchronized when the formula result changes.
// Use Cases: Visual progress bar made of repeated symbols that updates with a cell value. | Dynamic label that reflects any REPT formula for pattern previews or printable templates. | Automated report element where a shape mirrors a cell‑based text pattern without manual refresh.
// AI Prompts: Generate C# code to add a rectangle shape, link it to a cell containing a REPT formula, and refresh the shape's displayed text using Aspose.Cells. | Explain the two boolean arguments of SetLinkedCell and how UpdateSelectedValue synchronizes the shape with the linked cell. | Show how to modify the REPT formula after linking and ensure the shape updates automatically without recreating it.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, set a REPT("*",10) formula in cell A1, add a rectangle shape, link the shape to the formula cell using SetLinkedCell, refresh the displayed value, and save the file using Aspose.Cells for C#.
class ShapeLinkedToReptCell
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell A1 that repeats the character "*"
            // The REPT function repeats a text a given number of times
            sheet.Cells["A1"].Formula = "=REPT(\"*\", 10)";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 5, 5, 50, 200);

            // Link the shape to the cell containing the REPT formula (A1)
            // The two boolean parameters indicate whether to update the shape's value and whether to refresh the linked cell
            shape.SetLinkedCell("A1", true, true);

            // Update the shape so it displays the current value of the linked cell
            shape.UpdateSelectedValue();

            // Define output file path
            string outputPath = "ShapeLinkedToReptCell.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
