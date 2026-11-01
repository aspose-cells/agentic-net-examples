// Title: C# – Resize an Aspose.Cells shape proportionally to match a specific cell range while preserving aspect ratio
// Common Searches: how to proportionally resize a shape to a cell range using Aspose.Cells in C# | Aspose.Cells calculate pixel width and height of a worksheet range for shape scaling | C# set shape size to fit Excel cells without distortion with Aspose.Cells | place a shape at the top‑left corner of a target range in Aspose.Cells C# | preserve aspect ratio when fitting a picture into a range of Excel cells using Aspose
// Tags: proportional shape resizing Aspose.Cells | fit shape to Excel cell range C# | preserve aspect ratio shape Aspose.Cells | worksheet range pixel dimension calculation C# | shape upper left placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

// The example loads a workbook, defines a target cell range, computes its total pixel width and height, retrieves a shape by name, calculates a uniform scaling factor using the smaller of the width and height ratios to keep the original aspect ratio, resizes the shape accordingly, aligns its upper‑left corner with the first cell of the range, and saves the updated file.
class ShapeResizer
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell range (e.g., B2:D10)
            AsposeRange targetRange = sheet.Cells.CreateRange("B2:D10");

            // Calculate the pixel dimensions of the target range
            double targetWidth = 0;
            for (int col = targetRange.FirstColumn; col < targetRange.FirstColumn + targetRange.ColumnCount; col++)
                targetWidth += sheet.Cells.GetColumnWidthPixel(col);

            double targetHeight = 0;
            for (int row = targetRange.FirstRow; row < targetRange.FirstRow + targetRange.RowCount; row++)
                targetHeight += sheet.Cells.GetRowHeightPixel(row);

            // Retrieve the shape to be resized (by name)
            Shape shape = sheet.Shapes["MyShape"];
            if (shape == null)
                throw new InvalidOperationException("Shape named 'MyShape' was not found in the worksheet.");

            // Original dimensions of the shape in pixels
            double shapeWidth = shape.Width;
            double shapeHeight = shape.Height;

            // Compute scaling factor while preserving aspect ratio
            double widthScale = targetWidth / shapeWidth;
            double heightScale = targetHeight / shapeHeight;
            double scale = Math.Min(widthScale, heightScale); // choose the smaller factor

            // Apply proportional resizing (cast to int because Width/Height are integer pixels)
            shape.Width = (int)(shapeWidth * scale);
            shape.Height = (int)(shapeHeight * scale);

            // Position the shape at the top‑left corner of the target range
            shape.UpperLeftRow = targetRange.FirstRow;
            shape.UpperLeftColumn = targetRange.FirstColumn;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
