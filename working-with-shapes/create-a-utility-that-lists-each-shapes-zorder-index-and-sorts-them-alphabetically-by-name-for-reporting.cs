// Title: C# utility with Aspose.Cells to list each worksheet shape’s name and Z‑order index, then sort alphabetically
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, iterates over Worksheet.Shapes, captures Shape.Name and Shape.ZOrderPosition, orders the results by name, and writes a formatted console table. | Extend the shape‑listing program to also include Shape.Type in the output and export the sorted data to a CSV file using Aspose.Cells. | Add robust error handling so the utility checks for a missing input file, handles worksheets without shapes, and logs any exceptions to a log file.
// Common Searches: how to retrieve ZOrderPosition of shapes in an Excel file using Aspose.Cells C# | C# Aspose.Cells list all shapes on a worksheet and sort by shape name | example code to print shape name and Z‑order index from an .xlsx workbook | Aspose.Cells shape enumeration and alphabetical sorting tutorial | C# console app to report Excel shape metadata with Aspose.Cells
// Tags: Aspose.Cells enumerate worksheet shapes C# | retrieve shape ZOrderPosition Aspose.Cells | sort shape list alphabetically Aspose.Cells | C# console reporting of Excel shape metadata | export shape information to CSV Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeZOrderReporter
{
    // Loads an Excel workbook, iterates over all shapes on the first worksheet, collects each shape's Name and ZOrderPosition, sorts the collection alphabetically by name, and prints a table of shape names with their Z‑order indices.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify the desired one)
                Worksheet worksheet = workbook.Worksheets[0];

                // Collection to hold shape name and its Z‑order index
                List<(string Name, int ZOrder)> shapeInfo = new List<(string, int)>();

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Add shape name and Z‑order position to the list
                    shapeInfo.Add((shape.Name, shape.ZOrderPosition));
                }

                // Sort the list alphabetically by shape name
                var sortedShapes = shapeInfo.OrderBy(s => s.Name).ToList();

                // Output the sorted shape information
                Console.WriteLine("Shape Name\tZ‑Order Index");
                Console.WriteLine("-------------------------------");
                foreach (var s in sortedShapes)
                {
                    Console.WriteLine($"{s.Name}\t{s.ZOrder}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
