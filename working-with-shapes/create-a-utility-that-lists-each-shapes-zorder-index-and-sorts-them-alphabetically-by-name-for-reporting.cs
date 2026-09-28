// Title: Create a C# utility with Aspose.Cells to list each shape’s name and Z‑order index, then output an alphabetically sorted report
// AI Prompts: Load an .xlsx workbook with Aspose.Cells, iterate over Worksheet.Shapes, capture each shape’s Name (or generate a placeholder) and its ZOrderPosition, sort the collection by name, and print a tab‑delimited report. | Write C# code that reads the first worksheet, extracts shape names and Z‑order indices, orders the results alphabetically (case‑insensitive), and displays them in the console. | Extend the program to write the sorted shape name and Z‑order data to a CSV file using standard .NET I/O alongside Aspose.Cells.
// Common Searches: aspnet retrieve ZOrderPosition of shapes in an Excel file using Aspose.Cells | c# list all shapes on a worksheet and sort by name with Aspose.Cells | generate a shape name and Z‑order report from an .xlsx workbook in C# | aspose.cells get shape name placeholder when name is empty | output shape Z‑order index to console in C# Aspose.Cells example
// Tags: Aspose.Cells retrieve shape ZOrderPosition | C# list worksheet shapes alphabetically | shape name placeholder Aspose.Cells | sorted shape report .xlsx | console output shape Z-order Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The C# program loads an Excel workbook with Aspose.Cells, iterates through all shapes on the first worksheet, records each shape's name (or a generated placeholder) and its ZOrderPosition, sorts the entries alphabetically by name, and prints a tab‑delimited report of shape names and their Z‑order indices.
class ShapeZOrderReporter
{
    static void Main(string[] args)
    {
        // Path to the Excel file to be processed
        string filePath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File '{filePath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook contains no worksheets.");
                return;
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Collect shape name and its Z‑order index
            List<(string Name, int ZOrder)> shapes = new List<(string, int)>();

            // Iterate through all shapes on the worksheet using index for placeholder names
            for (int i = 0; i < sheet.Shapes.Count; i++)
            {
                Shape shape = sheet.Shapes[i];

                // Use the shape's Name property; if not set, create a placeholder using its index
                string name = string.IsNullOrEmpty(shape.Name) ? $"Shape_{i}" : shape.Name;

                // Z‑order index (higher value means the shape is in front)
                int zOrder = shape.ZOrderPosition;

                shapes.Add((name, zOrder));
            }

            // Sort the list alphabetically by shape name (case‑insensitive)
            shapes.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));

            // Output the sorted report
            Console.WriteLine("Shape Name\tZ-Order Index");
            Console.WriteLine("-------------------------------");
            foreach (var item in shapes)
            {
                Console.WriteLine($"{item.Name}\t{item.ZOrder}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
