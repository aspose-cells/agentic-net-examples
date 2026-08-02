// Title: Log Shape Adjustment Guide Changes in Aspose.Cells for .NET
// Description: Creates a workbook, adds a RightArrowCallout auto shape, initializes its ShapeGuideCollection, writes the original guide values to a text file, updates each guide, records every old‑to‑new change with timestamps, and saves both the workbook and the audit log.
// Keywords: Aspose.Cells | C# shape adjustment logging | audit shape guide changes | auto shape geometry tracking | Excel workbook audit log | ShapeGuideCollection write to file
// Common Searches: Aspose.Cells log shape guide changes | audit shape adjustments .NET | write shape geometry changes to text file | track auto shape modifications in Excel | record shape adjustment values Aspose.Cells
// Developer Intent: Capture every modification to a shape’s adjustment guides and store the details in a persistent text log for compliance or debugging.
// Use Cases: Generate an initial snapshot of all adjustment guides after shape creation. | Log each guide update during runtime, showing previous and new values. | Maintain a searchable audit trail that can be reviewed after the workbook is saved.
// AI Prompts: Create C# code that logs shape guide changes to a CSV file with columns for timestamp, shape ID, guide name, old value, and new value. | Provide a reusable method that accepts any Aspose.Cells shape, records its adjustment modifications, and appends entries to a log file. | Explain how to aggregate guide change logs from multiple worksheets into a single summary report.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeAdjustmentAudit
{
    // Creates a workbook, adds a RightArrowCallout auto shape, initializes its ShapeGuideCollection, writes the original guide values to a text file, updates each guide, records every old‑to‑new change with timestamps, and saves both the workbook and the audit log.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path for the audit log file
                string logPath = "ShapeAdjustmentsLog.txt";

                // Ensure the log file is empty at start
                File.WriteAllText(logPath, string.Empty);

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add an auto shape that supports adjustment guides (e.g., RightArrowCallout)
                Shape shape = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.RightArrowCallout, // shape type
                    2, 0,   // upper left row, column
                    2, 0,   // upper left row offset, column offset
                    200, 150); // width, height

                // Access the geometry adjustment collection
                ShapeGuideCollection guides = shape.Geometry.ShapeAdjustValues;

                // Add some initial adjustment guides
                guides.Add("adj1", 20.0);
                guides.Add("adj2", 30.0);
                guides.Add("adj3", 40.0);

                // Log initial adjustment values
                using (StreamWriter logWriter = new StreamWriter(logPath, true))
                {
                    logWriter.WriteLine($"Timestamp: {DateTime.Now}");
                    logWriter.WriteLine($"Shape ID: {shape.Name}");
                    logWriter.WriteLine("Initial adjustment values:");
                    for (int i = 0; i < guides.Count; i++)
                    {
                        ShapeGuide guide = guides[i];
                        logWriter.WriteLine($"  Guide{i + 1} = {guide.Value}");
                    }
                    logWriter.WriteLine();
                }

                // Example modification: change each guide value and log the change
                using (StreamWriter logWriter = new StreamWriter(logPath, true))
                {
                    logWriter.WriteLine($"Timestamp: {DateTime.Now}");
                    logWriter.WriteLine($"Shape ID: {shape.Name}");
                    logWriter.WriteLine("Modification of adjustment values:");

                    for (int i = 0; i < guides.Count; i++)
                    {
                        ShapeGuide guide = guides[i];
                        double oldValue = guide.Value;
                        double newValue = oldValue + 5.0; // arbitrary change

                        // Apply the new value
                        guide.Value = newValue;

                        // Log the change
                        logWriter.WriteLine($"  Guide{i + 1}: {oldValue} -> {newValue}");
                    }
                    logWriter.WriteLine();
                }

                // Save the workbook with the modified shape
                string outputPath = "ShapeAdjustmentAuditDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'. Audit log written to '{logPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
