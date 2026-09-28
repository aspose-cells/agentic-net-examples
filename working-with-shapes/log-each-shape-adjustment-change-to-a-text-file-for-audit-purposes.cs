// Title: Append timestamped audit entries for each shape width change while modifying shapes in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loops through all shapes on a worksheet with Aspose.Cells, increases each shape's width, and appends a UTC timestamped log line containing shape name, type, old width and new width to a text file. | Adapt the example to record height adjustments instead of width, preserving the same logging format and error‑handling pattern. | Extend the program to create a separate error‑log file that captures shape‑specific exceptions without stopping the processing of remaining shapes.
// Common Searches: Aspose.Cells C# how to log shape width changes to a file | record shape property modifications in Excel using Aspose.Cells .NET | append audit entries for shape adjustments with timestamp in Aspose.Cells | iterate worksheet shapes and save change history to text file in C#
// Tags: Aspose.Cells shape width audit logging | C# record shape property changes | timestamped shape adjustment log Aspose | append shape change entries to text file | error handling during shape iteration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for Shape class

namespace ShapeAdjustmentAudit
{
    // The sample loads an Excel workbook (creating a new one if the input file is missing), accesses the first worksheet, and iterates through every shape. For each shape it increases the Width by 5 points, writes a UTC‑timestamped line with the shape name, type, old width and new width to an audit log file (appending), and logs any shape‑specific exceptions. After processing, the workbook is saved to an output file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load workbook safely – create a new one if the input file is missing
                Workbook workbook;
                const string inputPath = "input.xlsx";
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // empty workbook
                }

                // Access the first worksheet (adjust as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare the audit log file (append mode)
                const string logPath = "shape_adjustments_log.txt";
                using (StreamWriter logWriter = new StreamWriter(logPath, true))
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        try
                        {
                            // Aspose.Cells does not expose shape adjustments directly.
                            // As an example, we modify the shape's width by a small increment
                            // and log the change. Adjust this logic as needed for your scenario.

                            int oldWidth = shape.Width;
                            int newWidth = oldWidth + 5; // increase width by 5 points
                            shape.Width = newWidth;

                            logWriter.WriteLine($"{DateTime.UtcNow:u} | Shape: {shape.Name} | Type: {shape.Type} | Width changed from {oldWidth} to {newWidth}");
                        }
                        catch (Exception shapeEx)
                        {
                            // Log shape‑specific errors without stopping the whole process
                            logWriter.WriteLine($"{DateTime.UtcNow:u} | Shape: {shape.Name} | Error: {shapeEx.Message}");
                        }
                    }
                }

                // Save workbook safely – overwrite if it exists
                const string outputPath = "output.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Log unexpected errors to console (or a separate error log)
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
