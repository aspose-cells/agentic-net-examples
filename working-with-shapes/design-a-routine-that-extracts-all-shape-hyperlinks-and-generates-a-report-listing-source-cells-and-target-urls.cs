// Title: Extract all shape hyperlinks from an Excel workbook and create a summary worksheet with source cells and target URLs using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method with Aspose.Cells that iterates through every worksheet, identifies shapes that have a Hyperlink, captures the worksheet name, shape name, the anchored cell address, and the hyperlink URL, then writes these details to a new worksheet named ShapeHyperlinkReport. | Create a C# console application that takes an input .xlsx file and an output path, extracts shape hyperlink data using Aspose.Cells, builds a report sheet with columns Worksheet, Shape Name, Source Cell, Target URL, and saves the modified workbook.
// Common Searches: how to list shape hyperlinks and their anchor cells in an Excel file using Aspose.Cells C# | Aspose.Cells C# extract hyperlink address from shapes and generate a report sheet | C# code to create a summary worksheet of all shape hyperlinks in a workbook with Aspose.Cells | retrieve source cell of a shape with hyperlink in .xlsx using Aspose.Cells for .NET | generate Excel report of shape hyperlink URLs and anchor locations with Aspose.Cells
// Tags: Aspose.Cells extract shape hyperlinks | C# generate hyperlink report worksheet | shape anchor cell address Aspose.Cells | Excel workbook shape hyperlink summary | Aspose.Cells iterate worksheets shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeHyperlinkExtractorApp
{
    // The example loads an Excel workbook, adds a new worksheet called "ShapeHyperlinkReport", scans each worksheet for shapes that contain a Hyperlink, records the worksheet name, shape name, the cell where the shape is anchored, and the hyperlink URL, writes this information into the report sheet, auto‑fits the columns, and saves the workbook to the specified output file.
    public class ShapeHyperlinkExtractor
    {
        /// <param name="inputFile">Path to the source workbook.</param>
        /// <param name="outputFile">Path where the workbook with the report will be saved.</param>
        public static void ExtractShapeHyperlinks(string inputFile, string outputFile)
        {
            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Add a new worksheet for the report
                int reportIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIndex];
                reportSheet.Name = "ShapeHyperlinkReport";

                // Write header row
                Cells reportCells = reportSheet.Cells;
                reportCells["A1"].PutValue("Worksheet");
                reportCells["B1"].PutValue("Shape Name");
                reportCells["C1"].PutValue("Source Cell");
                reportCells["D1"].PutValue("Target URL");

                int reportRow = 1; // zero‑based index; start after header

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Skip the report sheet itself
                    if (sheet == reportSheet) continue;

                    // Iterate through all shapes in the current worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check if the shape has a hyperlink defined
                        if (shape.Hyperlink != null && !string.IsNullOrEmpty(shape.Hyperlink.Address))
                        {
                            // Determine the source cell address where the shape is anchored
                            int row = shape.UpperLeftRow;
                            int col = shape.UpperLeftColumn;
                            string sourceCell = sheet.Cells[row, col].Name; // e.g., "B5"

                            // Write the information into the report sheet
                            reportCells[reportRow, 0].PutValue(sheet.Name);               // Worksheet
                            reportCells[reportRow, 1].PutValue(shape.Name);               // Shape Name
                            reportCells[reportRow, 2].PutValue(sourceCell);               // Source Cell
                            reportCells[reportRow, 3].PutValue(shape.Hyperlink.Address); // Target URL

                            reportRow++;
                        }
                    }
                }

                // Auto‑fit columns for better readability
                reportSheet.AutoFitColumns();

                // Save the workbook with the report
                workbook.Save(outputFile);
                Console.WriteLine($"Report saved to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – adjust paths as needed
            string inputPath = "Input.xlsx";
            string outputPath = "Output_With_Report.xlsx";

            ShapeHyperlinkExtractor.ExtractShapeHyperlinks(inputPath, outputPath);
        }
    }
}
