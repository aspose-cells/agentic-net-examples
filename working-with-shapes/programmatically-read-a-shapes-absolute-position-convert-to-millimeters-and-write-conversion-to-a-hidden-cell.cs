using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapePositionToHiddenCell
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Work with the first shape
            Shape shape = sheet.Shapes[0];

            // Get the actual box of the shape (values are in points)
            // box[0] = left (X), box[1] = top (Y), box[2] = width, box[3] = height
            float[] box = shape.GetActualBox();

            // Conversion factor: 1 point = 0.352777 mm
            const double pointToMm = 0.352777;

            double leftMm = box[0] * pointToMm;
            double topMm = box[1] * pointToMm;

            // Prepare the text to write (e.g., "Left: xx mm, Top: yy mm")
            string positionInfo = $"Left: {leftMm:F2} mm, Top: {topMm:F2} mm";

            // Write the conversion result to a hidden cell (e.g., Z1)
            Cell hiddenCell = sheet.Cells["Z1"];
            hiddenCell.PutValue(positionInfo);

            // Hide the column containing the cell (column Z -> index 25, zero‑based)
            int columnIndex = CellsHelper.ColumnNameToIndex("Z");
            sheet.Cells.HideColumn(columnIndex);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}