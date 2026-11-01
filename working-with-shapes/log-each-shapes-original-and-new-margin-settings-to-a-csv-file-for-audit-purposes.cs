// Title: Log original and updated shape margins to a CSV file while adjusting them with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook using Aspose.Cells, iterates through all shapes on the first worksheet, captures each shape's Top, Left, Bottom, and Right values, applies a configurable margin offset, and writes both the original and new coordinates to a CSV audit file. | Create a reusable C# method called AuditShapeMargins that takes the input workbook path, output workbook path, CSV log path, and a margin offset, updates each shape's position and size accordingly, saves the modified workbook, and returns the path of the generated CSV report.
// Common Searches: aspocells c# export shape coordinates to csv after resizing | how to audit shape margin changes in an Excel file using Aspose.Cells | log original and new shape positions in .NET Excel processing | record shape top left bottom right values before and after modification with Aspose.Cells | generate CSV report of shape margin adjustments in C#
// Tags: Aspose.Cells shape margin audit CSV | C# log shape coordinates Excel | adjust shape margins Aspose.Cells | export shape position to CSV .NET | shape margin offset Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates over each shape on the first worksheet, records the original Top/Left/Bottom/Right coordinates, adds a margin offset, writes both original and updated values to a CSV audit file, and saves the modified workbook.
class ShapeMarginAudit
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "Input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare CSV file for logging
            string csvPath = "ShapeMarginsAudit.csv";
            using (StreamWriter csvWriter = new StreamWriter(csvPath, false))
            {
                // Write CSV header
                csvWriter.WriteLine("ShapeName,OriginalTop,OriginalBottom,OriginalLeft,OriginalRight,NewTop,NewBottom,NewLeft,NewRight");

                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Capture original margin-like properties (using Top/Height/Left/Width as proxies)
                    double originalTop = shape.Top;
                    double originalLeft = shape.Left;
                    double originalBottom = shape.Top + shape.Height;
                    double originalRight = shape.Left + shape.Width;

                    // Example modification: increase each margin by 5 points
                    shape.Top += 5;
                    shape.Left += 5;
                    shape.Height += 5;
                    shape.Width += 5;

                    // Capture new values after modification
                    double newTop = shape.Top;
                    double newLeft = shape.Left;
                    double newBottom = shape.Top + shape.Height;
                    double newRight = shape.Left + shape.Width;

                    // Write the audit record to CSV
                    csvWriter.WriteLine($"{shape.Name},{originalTop},{originalBottom},{originalLeft},{originalRight},{newTop},{newBottom},{newLeft},{newRight}");
                }
            }

            // Save the modified workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to '{outputPath}'. Audit log saved to '{csvPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
