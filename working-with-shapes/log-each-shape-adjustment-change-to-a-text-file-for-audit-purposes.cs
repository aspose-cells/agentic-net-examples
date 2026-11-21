// Title: How to log shape position and size adjustments to a text file while modifying Excel shapes using Aspose.Cells for .NET
// Common Searches: C# code to write Aspose.Cells shape coordinate changes to a log file | record Excel shape size modifications with Aspose.Cells .NET example | audit shape adjustments in a workbook using Aspose.Cells and StreamWriter | log original and updated shape dimensions when moving shapes in Aspose.Cells | save shape adjustment details to a text file during Excel processing in C#
// Tags: Aspose.Cells shape coordinate logging | Excel shape dimension audit .NET | write shape adjustment log file C# | modify shape position Aspose.Cells | shape property change tracking Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file with Aspose.Cells, iterates through every worksheet and its shapes, writes each shape’s original Left, Top, Width and Height to a log file, applies a sample offset to those properties, records the adjusted values with a UTC timestamp, and saves the workbook to a new file.
class ShapeAdjustmentLogger
{
    static void Main()
    {
        // Define file paths
        string inputPath = @"C:\Input\Sample.xlsx";
        string logPath = @"C:\Logs\ShapeAdjustments.log";
        string outputPath = @"C:\Output\Sample_Adjusted.xlsx";

        try
        {
            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Ensure directories for log and output exist
            Directory.CreateDirectory(Path.GetDirectoryName(logPath));
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Open log file (overwrites if exists)
            using (StreamWriter logWriter = new StreamWriter(logPath, false))
            {
                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Log original position and size
                        logWriter.WriteLine($"Worksheet: {sheet.Name}, Shape: {shape.Name}");
                        logWriter.WriteLine($"  Original - Left: {shape.Left}, Top: {shape.Top}, Width: {shape.Width}, Height: {shape.Height}");

                        // Adjust shape (example logic)
                        shape.Left += 10;
                        shape.Top += 5;
                        shape.Width += 20;
                        shape.Height += 20;

                        // Log adjusted position and size
                        logWriter.WriteLine($"  Adjusted - Left: {shape.Left}, Top: {shape.Top}, Width: {shape.Width}, Height: {shape.Height}");
                        logWriter.WriteLine($"  Adjustment performed at: {DateTime.UtcNow:O}");
                        logWriter.WriteLine(); // Blank line for readability
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
            Console.WriteLine($"Log written to: {logPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
