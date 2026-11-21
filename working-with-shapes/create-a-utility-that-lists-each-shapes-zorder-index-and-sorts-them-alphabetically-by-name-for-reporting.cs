// Title: C# program to list each Excel shape’s Z‑order index and output an alphabetically sorted report using Aspose.Cells
// Common Searches: how to get ZOrderPosition of shapes in an Excel file with Aspose.Cells C# | list all shapes in a workbook and sort them by name using Aspose.Cells .NET | generate a shape Z‑order report from an .xlsx file in C# | retrieve shape names and Z‑order indices across worksheets with Aspose.Cells | C# console application to output Excel shape Z‑order sorted alphabetically
// Tags: Aspose.Cells shape ZOrderPosition extraction | C# enumerate Excel shapes across worksheets | alphabetical sorting of shape names Aspose.Cells | console reporting of shape Z‑order indices | placeholder naming for unnamed shapes Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeZOrderReporter
{
    // The utility loads an .xlsx workbook with Aspose.Cells, iterates through every worksheet’s ShapeCollection, captures each shape’s Name (or generates a placeholder when the name is empty) and its ZOrderPosition, stores these pairs, sorts them alphabetically ignoring case, and writes the name‑and‑Z‑order index report to the console (with optional file output).
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be processed
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(inputPath);

                // List to hold shape information (Name and Z‑order index)
                List<(string Name, int ZOrder)> shapeInfo = new List<(string, int)>();

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the collection of shapes on the current worksheet
                    ShapeCollection shapes = sheet.Shapes;

                    // Gather name and Z‑order index for each shape
                    foreach (Shape shape in shapes)
                    {
                        // Use shape name if available; otherwise create a placeholder using Z‑order position
                        string shapeName = string.IsNullOrEmpty(shape.Name)
                            ? $"Unnamed_{shape.ZOrderPosition}"
                            : shape.Name;

                        int zOrder = shape.ZOrderPosition;

                        shapeInfo.Add((shapeName, zOrder));
                    }
                }

                // Sort the collected shapes alphabetically by name (case‑insensitive)
                var sortedShapes = shapeInfo
                    .OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // Output the report to the console
                Console.WriteLine("Shape Z‑Order Report (sorted by name):");
                Console.WriteLine("--------------------------------------");
                foreach (var item in sortedShapes)
                {
                    Console.WriteLine($"Name: {item.Name}, Z‑Order Index: {item.ZOrder}");
                }

                // Optional: write the report to a text file
                // File.WriteAllLines("ShapeReport.txt", sortedShapes.Select(s => $"Name: {s.Name}, Z‑Order Index: {s.ZOrder}"));
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
