// Title: Cast SmartArt conversion result to GroupShape and verify its type – Aspose.Cells for .NET (C#)
// Description: This example shows how to create a workbook, add a rectangle placeholder, convert it with GetResultOfSmartArt, cast the returned object to GroupShape, confirm the cast succeeded, optionally adjust the group's position, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells GetResultOfSmartArt | GroupShape cast C# | SmartArt to GroupShape conversion | verify shape type Aspose.Cells | Aspose.Cells shape manipulation | C# Aspose.Cells GroupShape example | SmartArt conversion validation
// Common Searches: cast GetResultOfSmartArt to GroupShape Aspose.Cells | check if SmartArt conversion returns GroupShape C# | how to verify GroupShape after SmartArt conversion | Aspose.Cells modify GroupShape position | GetResultOfSmartArt returns null handling
// Developer Intent: The developer needs to cast the result of GetResultOfSmartArt to a GroupShape object and ensure the cast is valid before manipulating the shape.
// Use Cases: Validate that a SmartArt placeholder was successfully converted to a GroupShape before applying layout changes. | Adjust the Left and Top properties of the GroupShape after conversion. | Conditionally save the workbook only when the GroupShape conversion is non‑null and correctly typed.
// AI Prompts: Generate C# code that casts shape.GetResultOfSmartArt() to GroupShape and logs whether the cast succeeded. | Provide an example that iterates through all child shapes inside the GroupShape returned by GetResultOfSmartArt. | Show how to handle a null result from GetResultOfSmartArt when the source shape is not SmartArt.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example shows how to create a workbook, add a rectangle placeholder, convert it with GetResultOfSmartArt, cast the returned object to GroupShape, confirm the cast succeeded, optionally adjust the group's position, and save the file using Aspose.Cells for .NET.
    public class CastSmartArtResultDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape (placeholder for SmartArt)
                Shape shape = sheet.Shapes.AddRectangle(0, 0, 0, 0, 200, 200);

                // Convert SmartArt to grouped shapes (returns GroupShape if shape is SmartArt)
                GroupShape result = shape.GetResultOfSmartArt();

                // Verify that the conversion result is a GroupShape
                if (result != null && result is GroupShape)
                {
                    Console.WriteLine("Conversion succeeded: result is a GroupShape.");
                    // Example: modify the group shape
                    result.Left = 100;
                    result.Top = 50;
                }
                else
                {
                    Console.WriteLine("Conversion did not produce a GroupShape (result is null or not a GroupShape).");
                }

                // Save the workbook
                workbook.Save("CastSmartArtResultDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
