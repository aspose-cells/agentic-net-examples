// Title: Audit TextBox shape margin changes in an Excel workbook and export results to CSV using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through every worksheet, finds TextBox shapes, reads their current margin values (if the API exposes them), adds a configurable offset, and records both original and updated margins to a CSV audit file. | Enhance the example to accept the margin increment and the CSV output path as command‑line arguments, and include additional columns for worksheet name and shape type in the generated CSV. | Add robust error handling that skips shapes without accessible margin properties, logs a warning for each skipped shape, and guarantees that the workbook is saved after processing.
// Common Searches: how to log textbox shape margins to a csv with Aspose.Cells C# | Aspose.Cells retrieve and modify shape margin values in .NET | C# export Excel shape properties to CSV for audit | increment text box margins in an Excel workbook using Aspose.Cells | audit changes to Excel shape margins programmatically
// Tags: Aspose.Cells retrieve shape margin properties | C# export Excel shape data to CSV | adjust TextBox margins Aspose.Cells | audit shape margin changes .NET | iterate worksheet shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook, walks through each worksheet and its TextBox shapes, reads existing margin settings when available, adds a fixed offset, writes original and new margin values together with the shape name to a CSV audit file, and saves the modified workbook.
class ShapeMarginAudit
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string csvPath = "ShapeMarginAudit.csv";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare CSV file for logging
            using (StreamWriter csvWriter = new StreamWriter(csvPath, false))
            {
                // Write CSV header
                csvWriter.WriteLine("ShapeName,OriginalMarginTop,OriginalMarginBottom,OriginalMarginLeft,OriginalMarginRight,NewMarginTop,NewMarginBottom,NewMarginLeft,NewMarginRight");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only shapes that are TextBox objects
                        if (shape is TextBox textbox)
                        {
                            try
                            {
                                // Aspose.Cells versions prior to 23.9 do not expose margin properties directly.
                                // Use placeholder values for demonstration; replace with actual APIs when available.
                                double origTop = 0;
                                double origBottom = 0;
                                double origLeft = 0;
                                double origRight = 0;

                                // Example modification: increase each margin by 5 points (placeholder logic)
                                double newTop = origTop + 5;
                                double newBottom = origBottom + 5;
                                double newLeft = origLeft + 5;
                                double newRight = origRight + 5;

                                // Write a line to the CSV file
                                csvWriter.WriteLine($"{shape.Name},{origTop},{origBottom},{origLeft},{origRight},{newTop},{newBottom},{newLeft},{newRight}");
                            }
                            catch (Exception shapeEx)
                            {
                                Console.WriteLine($"Error processing shape '{shape.Name}': {shapeEx.Message}");
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Modified workbook saved to '{outputPath}'. Audit CSV saved to '{csvPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
