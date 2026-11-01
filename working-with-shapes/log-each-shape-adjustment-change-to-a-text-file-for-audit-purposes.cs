// Title: Log shape size changes to a text file while adjusting Excel shapes with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx workbook using Aspose.Cells, increases a shape's width and height, and appends the shape's name, type, position, and new dimensions to a timestamped log file. | Create a reusable C# method that takes any Aspose.Cells Shape object and a log file path, then writes the shape's current properties (name, type, left, top, width, height) together with a custom context label to the file. | Enhance the shape‑logging routine to catch I/O exceptions, record error details in a separate error log, and allow the workbook save operation to continue.
// Common Searches: how to record before and after shape dimensions in Aspose.Cells C# | Aspose.Cells log shape modifications to a text file example | C# audit changes to Excel drawing objects using Aspose.Cells | save shape property audit trail when resizing shapes in .xlsx with Aspose.Cells
// Tags: Aspose.Cells shape resize logging | C# audit Excel shape properties | log shape dimensions to text file | timestamped shape change audit Aspose.Cells | exception handling for shape audit logging

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeAdjustmentAudit
{
    // The sample loads an existing workbook, checks for shapes on the first worksheet, logs each shape's properties before and after resizing to a text file with timestamps, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";
                string logFilePath = "shape_audit_log.txt";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Example modification to a shape and audit logging
                if (sheet.Shapes.Count > 0)
                {
                    Shape shape = sheet.Shapes[0];

                    // Log shape state before modification
                    LogShapeChange(shape, "Before modification", logFilePath);

                    // Perform modifications that would constitute a shape adjustment
                    shape.Width += 10;   // Increase width
                    shape.Height += 5;   // Increase height

                    // Log shape state after modification
                    LogShapeChange(shape, "After modification", logFilePath);
                }
                else
                {
                    Console.WriteLine("No shapes found on the worksheet to modify.");
                }

                // Save the workbook after modifications
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void LogShapeChange(Shape shape, string context, string logFilePath)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {context}");
                    writer.WriteLine($"Shape '{shape.Name}' (Type: {shape.Type})");
                    writer.WriteLine($"   Position: Left={shape.Left}, Top={shape.Top}");
                    writer.WriteLine($"   Size: Width={shape.Width}, Height={shape.Height}");
                    writer.WriteLine(); // Blank line for readability
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }
    }
}
