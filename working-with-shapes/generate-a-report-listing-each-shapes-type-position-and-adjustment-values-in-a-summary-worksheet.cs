// Title: Create a summary worksheet that lists each shape’s type, coordinates, size, and adjustment values in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console application that opens an existing .xlsx workbook with Aspose.Cells, adds a new worksheet called "Summary", and records for every shape on each non‑summary sheet its name, type, top, left, width, height, and any adjustment values. | Enhance the shape‑reporting code to safely handle shapes without a name, exclude the summary sheet from the iteration, and log any shape‑processing errors without terminating the program. | After populating the summary data, apply AutoFit to all columns and save the modified workbook to a separate output file while preserving the original workbook.
// Common Searches: Aspose.Cells C# list all shapes in a workbook with their position and size | How to export shape type and coordinates to a new worksheet using Aspose.Cells .NET | C# generate shape report including adjustment values in Excel with Aspose.Cells | Skip a specific worksheet while iterating shapes in Aspose.Cells | AutoFit columns after writing data with Aspose.Cells C#
// Tags: Aspose.Cells extract shape properties C# | generate shape summary worksheet Aspose.Cells | list shape coordinates Aspose.Cells .NET | export shape adjustments Aspose.Cells | auto‑fit columns Aspose.Cells workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeReportGenerator
{
    // The program loads an existing workbook, adds a 'Summary' sheet, iterates all other worksheets, extracts each shape’s name, type, top, left, width, height (adjustments left blank), writes these details to the summary sheet, auto‑fits columns for readability, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Add a new worksheet for the summary report
                Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
                summarySheet.Name = "Summary";

                // Write header row
                Cells cells = summarySheet.Cells;
                cells[0, 0].PutValue("Worksheet");
                cells[0, 1].PutValue("Shape Name");
                cells[0, 2].PutValue("Shape Type");
                cells[0, 3].PutValue("Top");
                cells[0, 4].PutValue("Left");
                cells[0, 5].PutValue("Width");
                cells[0, 6].PutValue("Height");
                cells[0, 7].PutValue("Adjustments");

                int currentRow = 1; // Start after header

                // Iterate through all worksheets and their shapes
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the summary sheet itself to avoid self‑reference
                    if (ws.Name == summarySheet.Name) continue;

                    foreach (Shape shape in ws.Shapes)
                    {
                        try
                        {
                            // Shape name (if not set, use empty string)
                            string shapeName = string.IsNullOrEmpty(shape.Name) ? "" : shape.Name;

                            // Shape type as string
                            string shapeType = shape.Type.ToString();

                            // Position and size
                            double top = shape.Top;
                            double left = shape.Left;
                            double width = shape.Width;
                            double height = shape.Height;

                            // Adjustments are not directly exposed in Aspose.Cells Shape;
                            // leave empty or implement custom logic if needed.
                            string adjustments = "";

                            // Write data to the summary sheet
                            cells[currentRow, 0].PutValue(ws.Name);
                            cells[currentRow, 1].PutValue(shapeName);
                            cells[currentRow, 2].PutValue(shapeType);
                            cells[currentRow, 3].PutValue(top);
                            cells[currentRow, 4].PutValue(left);
                            cells[currentRow, 5].PutValue(width);
                            cells[currentRow, 6].PutValue(height);
                            cells[currentRow, 7].PutValue(adjustments);

                            currentRow++;
                        }
                        catch (Exception shapeEx)
                        {
                            Console.WriteLine($"Error processing shape in worksheet '{ws.Name}': {shapeEx.Message}");
                        }
                    }
                }

                // Auto‑fit columns for better readability
                summarySheet.AutoFitColumns();

                // Save the workbook with the new summary sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Summary report saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
