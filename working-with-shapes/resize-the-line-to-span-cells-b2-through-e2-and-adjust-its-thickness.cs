using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ResizeLineExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a line shape with temporary size (will be resized later)
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (px), width (px)
            LineShape line = worksheet.Shapes.AddLine(0, 0, 0, 0, 0, 0);

            // Position the line so that it starts at cell B2 and ends at cell E2
            // Rows and columns are zero‑based: B2 => row 1, column 1; E2 => row 1, column 4
            line.UpperLeftRow = 1;      // start row (B2)
            line.UpperLeftColumn = 1;   // start column (B2)
            line.LowerRightRow = 1;     // end row (E2)
            line.LowerRightColumn = 4;  // end column (E2)

            // Remove any pixel offsets so the line aligns exactly with the cell borders
            line.UpperDeltaX = 0;
            line.UpperDeltaY = 0;
            line.LowerDeltaX = 0;
            line.LowerDeltaY = 0;

            // Adjust the line thickness (weight) – using points for clarity
            line.Line.Weight = 2.5f;   // 2.5 points thick

            // Define output file path
            string outputPath = "ResizedLine.xlsx";

            // Ensure the directory exists (optional safety)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook with the resized line
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}