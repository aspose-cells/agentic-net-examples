// Title: Detect SmartArt Shapes in an Excel Workbook Using Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, loops through every worksheet and its ShapeCollection, uses the Shape.IsSmartArt property to identify SmartArt objects, logs worksheet name, shape index and shape name to the console, and optionally saves the workbook.
// Keywords: Aspose.Cells SmartArt detection | C# Shape.IsSmartArt | iterate Excel shapes .NET | list SmartArt objects | Excel shape collection Aspose | GitHub Aspose.Cells example | Answer Engine Optimization Excel SmartArt
// Common Searches: how to find SmartArt in Excel using Aspose.Cells C# | Aspose.Cells iterate shapes and detect SmartArt | Shape.IsSmartArt property example | list all SmartArt shapes in a workbook | C# code to enumerate SmartArt objects in Excel
// Developer Intent: Programmatically locate and enumerate every SmartArt shape across all worksheets in an Excel workbook.
// Use Cases: Generate an audit report of SmartArt objects with their sheet names and indices. | Apply batch formatting or replace SmartArt shapes automatically. | Extract SmartArt metadata for downstream processing such as image export or data analysis.
// AI Prompts: Write C# code that replaces each SmartArt shape in a workbook with a placeholder image using Aspose.Cells. | Create a method that returns a List<Shape> containing all SmartArt objects from a given worksheet. | Develop a script that extracts the text from every SmartArt shape and saves it to a CSV file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtDemo
{
    // Loads an Excel file, loops through every worksheet and its ShapeCollection, uses the Shape.IsSmartArt property to identify SmartArt objects, logs worksheet name, shape index and shape name to the console, and optionally saves the workbook.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes in the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Iterate through each shape
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Identify SmartArt objects using the IsSmartArt property
                    if (shape.IsSmartArt)
                    {
                        Console.WriteLine($"SmartArt found in worksheet \"{sheet.Name}\" at shape index {i} (Name: {shape.Name})");
                        // Additional processing for SmartArt can be placed here
                    }
                }
            }

            // Save the workbook (optional, if modifications were made)
            workbook.Save("output.xlsx");
        }
    }
}
