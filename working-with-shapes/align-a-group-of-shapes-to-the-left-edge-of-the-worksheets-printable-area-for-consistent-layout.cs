// Title: Align all worksheet shapes to the left edge of the printable area using Aspose.Cells for .NET (C#)
// AI Prompts: Move every shape on the first worksheet to the first printable column of the defined print area while preserving its row using Aspose.Cells in C#. | Set each shape's UpperLeftColumn to the leftmost column of the worksheet's print area based on PageSetup.PrintArea with Aspose.Cells API.
// Common Searches: asp.net aspose.cells align shapes to left margin of print area c# | how to set shape UpperLeftColumn to first printable column aspose.cells | move all worksheet drawings to left edge of defined print area using Aspose.Cells | c# adjust shape positions according to page setup print area aspose.cells
// Tags: align shapes left printable area Aspose.Cells | shape UpperLeftColumn positioning Aspose.Cells | worksheet print area column alignment C# | adjust shape column based on page setup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

// The example loads a workbook, determines the worksheet's print area (or falls back to the used range), calculates the first printable column, iterates through all shapes on the first worksheet, sets each shape's UpperLeftColumn to that column while keeping its original row, and saves the updated file.
class AlignShapesToPrintableArea
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the defined print area; fallback to the used range if not set.
            string printArea = sheet.PageSetup.PrintArea;
            if (string.IsNullOrEmpty(printArea))
            {
                AsposeRange used = sheet.Cells.MaxDisplayRange;
                // Convert row/column indices to Excel A1 notation.
                string firstCell = CellsHelper.CellIndexToName(used.FirstRow, used.FirstColumn);
                string lastCell = CellsHelper.CellIndexToName(
                    used.FirstRow + used.RowCount - 1,
                    used.FirstColumn + used.ColumnCount - 1);
                printArea = $"{firstCell}:{lastCell}";
            }

            // Create a Range object from the print area string.
            AsposeRange printRange = sheet.Cells.CreateRange(printArea);

            // Determine the first column index of the printable area (0‑based).
            int firstPrintableColumn = printRange.FirstColumn;

            // Align each shape to the left edge of the printable area.
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Preserve the shape's current row position.
                    int currentRow = shape.UpperLeftRow;

                    // Align shape to the first printable column with no horizontal offset.
                    shape.UpperLeftRow = currentRow;
                    shape.UpperLeftColumn = firstPrintableColumn;
                    // Offsets default to 0; no need to set UpperLeftColumnOffset if not supported.
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Failed to adjust shape '{shape.Name}': {exShape.Message}");
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
