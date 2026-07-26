// Title: Link a TextBox shape to a TEXTJOIN cell with Aspose.Cells for .NET
// Description: Creates a workbook, fills B1:B5 with fruit names, applies a TEXTJOIN formula in A1, adds a TextBox shape, links the shape to A1 so it shows the concatenated list, formats the text, and saves the file as LinkedShapeWithTextJoin.xlsx.
// Keywords: Aspose.Cells | .NET | C# | TextBox shape | SetLinkedCell | TEXTJOIN formula | Excel automation | dynamic list | linked shape | cell reference
// Common Searches: Aspose.Cells link shape to cell example | C# TEXTJOIN with Aspose.Cells | SetLinkedCell method usage | display formula result in TextBox shape | dynamic shape content Excel .NET
// Developer Intent: Show how to bind a TextBox shape to a cell that contains a TEXTJOIN formula so the shape updates automatically.
// Use Cases: Dashboard widgets that reflect combined values from a column. | Reports where a shape displays a merged list of items. | Interactive worksheets that auto‑refresh shape text when source data changes.
// AI Prompts: Generate C# code using Aspose.Cells to add a TextBox, link it to a TEXTJOIN cell, and set font properties. | Explain how to refresh a linked shape after expanding the TEXTJOIN range. | Describe the purpose of the two boolean flags in SetLinkedCell when linking to a formula cell.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// Creates a workbook, fills B1:B5 with fruit names, applies a TEXTJOIN formula in A1, adds a TextBox shape, links the shape to A1 so it shows the concatenated list, formats the text, and saves the file as LinkedShapeWithTextJoin.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column B (B1:B5)
            cells["B1"].PutValue("Apple");
            cells["B2"].PutValue("Banana");
            cells["B3"].PutValue("Cherry");
            cells["B4"].PutValue("Date");
            cells["B5"].PutValue("Elderberry");

            // Set a TEXTJOIN formula in cell A1 to concatenate the values in B1:B5
            // Formula: =TEXTJOIN(", ", TRUE, B1:B5)
            cells["A1"].Formula = "=TEXTJOIN(\", \", TRUE, B1:B5)";

            // Add a TextBox shape to the worksheet
            // Parameters: upper left row, upper left column, lower right row, lower right column, height, width
            int upperLeftRow = 7;
            int upperLeftColumn = 1;
            int lowerRightRow = 12;
            int lowerRightColumn = 5;
            int height = 100; // pixels
            int width = 200;  // pixels
            TextBox textBox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, height, width);

            // Link the TextBox to cell A1 so it displays the concatenated list
            // The two boolean parameters indicate whether to set the linked cell as a formula and whether to update the shape immediately
            textBox.SetLinkedCell("A1", true, true);

            // Adjust the shape appearance
            textBox.Text = ""; // Clear default text; linked cell value will be shown
            textBox.Font.Size = 12;
            textBox.Font.Color = Color.Blue;

            // Define output file path
            string outputPath = "LinkedShapeWithTextJoin.xlsx";

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
