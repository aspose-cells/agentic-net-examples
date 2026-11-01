// Title: How to retrieve all shapes from an Excel worksheet, sort them by Z‑order, and print the sorted list using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, extracts every Shape from the first worksheet, sorts the shapes by their ZOrderPosition property, and writes each shape's name and Z‑order index to the console. | Create a C# method using Aspose.Cells that returns a List<Shape> ordered by Z‑order for a given worksheet and show how to iterate the list to display shape names and Z‑order values.
// Common Searches: aspocells c# list worksheet shapes sorted by Z-order | how to get shape ZOrderPosition in Aspose.Cells .NET | sort Excel drawing objects by Z-index using Aspose.Cells | retrieve and display shape names and Z-order from an Excel file in C# | Aspose.Cells example for enumerating and ordering shapes on a worksheet
// Tags: Aspose.Cells shape enumeration C# | sort shapes by ZOrderPosition Aspose.Cells | list worksheet drawing objects .NET | retrieve shape names Excel Aspose.Cells | order Excel shapes by Z-index C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel file, accesses the first worksheet, collects all Shape objects, sorts them by their ZOrderPosition (ascending), and prints each shape's name and Z‑order index to the console, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main(string[] args)
    {
        // Path to the input workbook
        string filePath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve all shapes from the worksheet
            ShapeCollection shapeCollection = sheet.Shapes;

            // Copy shapes to a list for sorting
            List<Shape> shapes = new List<Shape>();
            foreach (Shape shape in shapeCollection)
            {
                shapes.Add(shape);
            }

            // Sort shapes by their Z-order position (ascending)
            List<Shape> sortedShapes = shapes.OrderBy(s => s.ZOrderPosition).ToList();

            // Output sorted list to console
            Console.WriteLine("Shapes sorted by Z-order:");
            foreach (Shape shape in sortedShapes)
            {
                Console.WriteLine($"Name: {shape.Name}, Z-Order: {shape.ZOrderPosition}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
