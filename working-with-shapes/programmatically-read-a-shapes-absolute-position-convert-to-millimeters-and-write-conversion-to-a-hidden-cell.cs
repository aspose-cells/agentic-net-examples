// Title: Aspose.Cells for .NET – Read Shape Position, Convert to Millimeters, and Save in a Hidden Cell (C#)
// Description: Loads a workbook, gets the first shape's absolute LeftCM and TopCM values, converts them to millimetres, writes the formatted result to a cell, hides the cell by applying a transparent font, and saves the file.
// Keywords: Aspose.Cells shape position | convert shape coordinates to mm | LeftCM TopCM Aspose.Cells | hidden cell transparent font | C# read shape location | Aspose.Cells write hidden value
// Common Searches: how to get shape left and top in Aspose.Cells | convert shape position from cm to mm in C# | store shape coordinates in a hidden Excel cell | hide cell text with transparent font Aspose.Cells
// Developer Intent: Retrieve a shape's absolute coordinates, transform them to millimetres, and record the values in a concealed worksheet cell.
// Use Cases: Audit exact placement of graphics without exposing data to end users. | Drive formulas or conditional formatting using hidden shape‑position values. | Generate documentation that logs shape locations for later analysis.
// AI Prompts: Write C# code that reads a shape's LeftCM and TopCM, converts the values to millimetres, and writes them to a hidden cell using Aspose.Cells. | Explain how to hide cell content by setting the font color to Transparent in Aspose.Cells for .NET. | Show how to loop through all shapes on a worksheet and store each shape's converted position in separate hidden cells.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, gets the first shape's absolute LeftCM and TopCM values, converts them to millimetres, writes the formatted result to a cell, hides the cell by applying a transparent font, and saves the file.
class ShapePositionToHiddenCell
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Process the first shape if any exist
            if (worksheet.Shapes.Count > 0)
            {
                try
                {
                    Shape shape = worksheet.Shapes[0];

                    // Get the shape's position in centimeters
                    double leftCm = shape.LeftCM;
                    double topCm = shape.TopCM;

                    // Convert centimeters to millimeters (1 cm = 10 mm)
                    double leftMm = leftCm * 10.0;
                    double topMm = topCm * 10.0;

                    // Write the converted values to a cell (e.g., B2)
                    Cell targetCell = worksheet.Cells["B2"];
                    targetCell.PutValue($"Left: {leftMm:F2} mm, Top: {topMm:F2} mm");

                    // Hide the cell's content by setting font color to Transparent
                    Style hiddenStyle = targetCell.GetStyle();
                    hiddenStyle.Font.Color = Color.Transparent;
                    targetCell.SetStyle(hiddenStyle);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing shape: {ex.Message}");
                }
            }

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
