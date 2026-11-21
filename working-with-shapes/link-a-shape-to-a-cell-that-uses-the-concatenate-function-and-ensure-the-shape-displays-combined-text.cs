// Title: Link a textbox shape to a CONCATENATE formula cell in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link textbox shape to cell formula result | display CONCATENATE formula output in an Excel shape using Aspose.Cells | bind shape text to A1 formula value with Aspose.Cells .NET | create a linked textbox that shows combined cell values in Aspose.Cells | set shape.Text to a cell reference after calculating formulas in Aspose.Cells | link an Excel shape to a cell containing CONCATENATE in C#
// Tags: textbox shape formula binding Aspose.Cells | link shape to cell value .NET | concatenate function shape text Aspose | Aspose.Cells linked shape example C# | Excel shape text from formula Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes "Hello" and "World" to B1 and C1, sets A1 to a CONCATENATE formula, adds a textbox shape, links the shape's text to cell A1 using a formula reference, forces formula calculation, and saves the file as LinkedShape.xlsx.
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

            // Populate cells B1 and C1 with sample text
            sheet.Cells["B1"].PutValue("Hello");
            sheet.Cells["C1"].PutValue("World");

            // Set A1 to a CONCATENATE formula that combines B1 and C1 with a space
            sheet.Cells["A1"].Formula = "=CONCATENATE(B1, \" \", C1)";

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, height, width
            int upperLeftRow = 2;          // Row index (0‑based)
            int upperLeftColumn = 0;       // Column index (0‑based)
            int upperLeftRowOffset = 0;    // Offset in pixels
            int upperLeftColumnOffset = 0; // Offset in pixels
            int height = 100;              // Height in points
            int width = 200;               // Width in points

            Shape shape = sheet.Shapes.AddTextBox(
                upperLeftRow,
                upperLeftColumn,
                upperLeftRowOffset,
                upperLeftColumnOffset,
                height,
                width);

            // Link the shape's text to cell A1 using a formula reference
            shape.Text = "=A1";

            // Force calculation of all formulas so the linked shape displays the result
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputPath = "LinkedShape.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
