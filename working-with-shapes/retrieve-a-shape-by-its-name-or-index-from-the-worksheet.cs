// Title: How to retrieve a worksheet shape by name or index using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that fetches a shape from a worksheet by its assigned name using Aspose.Cells. | Show a C# example that accesses a shape from the Shapes collection by zero‑based index with bounds checking in Aspose.Cells. | Provide a C# snippet that verifies a shape exists before reading its properties from an Excel workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# retrieve shape by name from worksheet | C# get shape at specific index using Aspose.Cells Shapes collection | Check if a shape exists in an Excel file before accessing it with Aspose.Cells | How to list and access shapes in a .xlsx workbook using Aspose.Cells for .NET
// Tags: shape retrieval by name Aspose.Cells | shape retrieval by index Aspose.Cells | worksheet shapes collection handling Aspose.Cells | verify shape existence Aspose.Cells | load workbook and access shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates loading a workbook, then retrieving a shape named "MyShape" and a shape at index 0 from the first worksheet using Aspose.Cells for .NET, with safety checks and console output.
class RetrieveShapeExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Retrieve shape by its name -----
            Shape shapeByName = sheet.Shapes["MyShape"];
            if (shapeByName != null)
            {
                Console.WriteLine($"Shape found by name: {shapeByName.Name}, Type: {shapeByName.Type}");
            }
            else
            {
                Console.WriteLine("No shape with the name 'MyShape' was found.");
            }

            // ----- Retrieve shape by its index -----
            int shapeIndex = 0; // example index
            if (shapeIndex >= 0 && shapeIndex < sheet.Shapes.Count)
            {
                Shape shapeByIndex = sheet.Shapes[shapeIndex];
                Console.WriteLine($"Shape found by index {shapeIndex}: {shapeByIndex.Name}, Type: {shapeByIndex.Type}");
            }
            else
            {
                Console.WriteLine($"Shape index {shapeIndex} is out of range. Total shapes: {sheet.Shapes.Count}");
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
