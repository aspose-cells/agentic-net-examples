// Title: Apply a 50% transparent reflection to a shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Set the Reflection.Transparency property of a worksheet shape to 0.5 with Aspose.Cells in C#. | Create a rectangle shape when none exist and apply a 50% transparent reflection using Shape.Reflection in Aspose.Cells. | Update an existing Excel shape's reflection effect to half‑transparent and save the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to make shape reflection 50% transparent | set reflection transparency on Excel shape using Aspose.Cells .NET | add rectangle shape and apply reflection effect in Aspose.Cells workbook | modify shape reflection property in C# Aspose.Cells example | Excel shape reflection settings Aspose.Cells API
// Tags: Aspose.Cells shape.Reflection.Transparency | C# add rectangle shape Aspose.Cells | apply half transparent reflection Aspose.Cells | modify existing shape properties Excel .NET | reflection effect on worksheet shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, accesses the first worksheet, retrieves the first shape or adds a rectangle if none exist, sets its Reflection.Transparency to 0.5 (50% transparent), and saves the modified file as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load existing workbook if it exists; otherwise create a new one
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the first shape or add a placeholder rectangle if none exist
                Shape shape;
                if (worksheet.Shapes.Count > 0)
                {
                    shape = worksheet.Shapes[0];
                }
                else
                {
                    // Add a rectangle shape as a placeholder
                    shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                }

                // Apply a reflection effect (type defaults to Reflection)
                shape.Reflection.Transparency = 0.5; // 50% transparent reflection

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
