// Title: Adding a rectangle shape linked to a summary cell on every worksheet with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add rectangle shape to each worksheet and link to summary sheet cell | Iterate through worksheets and insert shape with internal hyperlink using Aspose.Cells .NET | How to create a shape that opens a summary cell in an Aspose.Cells workbook | Skip a specific sheet while adding shapes in Aspose.Cells C# | Set shape text and line weight in Aspise.Cells for .NET
// Tags: rectangle shape insertion Aspose.Cells | shape internal hyperlink Aspose.Cells | exclude summary worksheet Aspose.Cells | shape text and line weight Aspose.Cells | worksheet loop Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, loops through all worksheets except the one named "Summary", adds a rectangular shape labeled "View Summary" to each sheet, assigns an internal hyperlink that points to cell A1 on the Summary sheet, optionally sets the shape's line weight, logs each operation, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the address of the summary cell to which all shapes will link
            string summarySheetName = "Summary";
            string summaryCellAddress = "A1";
            string summaryLink = $"{summarySheetName}!{summaryCellAddress}";

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Skip the summary sheet itself if it exists
                if (sheet.Name.Equals(summarySheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    // Add a rectangle shape to the worksheet
                    // Parameters: type, upper-left column, upper-left row, column offset, row offset, width, height
                    Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

                    // Set a simple text inside the shape
                    shape.Text = "View Summary";

                    // Link the shape to the summary cell (internal hyperlink)
                    shape.Hyperlink.Address = summaryLink;

                    // Optional: format the shape (e.g., line weight)
                    shape.Line.Weight = 1.0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add shape to worksheet '{sheet.Name}': {ex.Message}");
                }

                Console.WriteLine($"Added shape to worksheet '{sheet.Name}' linking to {summaryLink}.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
