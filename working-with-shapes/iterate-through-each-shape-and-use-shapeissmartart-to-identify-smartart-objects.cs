// Title: Identify and list SmartArt shapes in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, loops through every worksheet and each Shape, and prints the Name, UpperLeftRow and UpperLeftColumn of shapes where Shape.IsSmartArt is true. | Create a .NET console application that safely loads a workbook, checks for a missing input file, uses Shape.IsSmartArt to detect SmartArt objects, and logs their positions without altering the file. | Show how to catch and log exceptions while scanning an Excel workbook for SmartArt shapes with Aspose.Cells and output the results to the console.
// Common Searches: C# Aspose.Cells how to enumerate SmartArt objects in an Excel sheet | retrieve SmartArt shape coordinates using Shape.IsSmartArt property | list all SmartArt names and cell locations in a workbook with Aspose.Cells for .NET | detect SmartArt diagrams in Excel files programmatically using Aspose.Cells | example code to find SmartArt shapes across worksheets in C#
// Tags: enumerate SmartArt shapes Aspose.Cells | Shape.IsSmartArt detection C# | list SmartArt coordinates Excel | scan workbook shapes Aspose.Cells | handle missing Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing .xlsx workbook, iterates through each worksheet and its shapes, uses the Shape.IsSmartArt property to identify SmartArt objects, and writes each object's name and top‑left row/column to the console before optionally saving the workbook.
class SmartArtDetector
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Identify SmartArt objects using the IsSmartArt property
                    if (shape.IsSmartArt)
                    {
                        // Output the shape name and its position (upper‑left cell)
                        Console.WriteLine($"SmartArt found on sheet \"{sheet.Name}\":");
                        Console.WriteLine($"  Name: {shape.Name}");
                        Console.WriteLine($"  Top‑Left Row: {shape.UpperLeftRow}, Column: {shape.UpperLeftColumn}");
                    }
                }
            }

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
