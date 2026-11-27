// Title: How to recursively ungroup nested shape groups in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# code to ungroup all shape groups in an Excel worksheet with Aspose.Cells | remove nested group shapes from a workbook using Aspose.Cells .NET API | iterate through worksheet.Shapes and flatten grouped shapes recursively Aspose.Cells | Aspose.Cells ungroup GroupShape objects until only individual shapes remain | programmatically ungroup Excel shapes hierarchy in C#
// Tags: Aspose.Cells recursive shape ungrouping | GroupShape.Ungroup method C# | flatten Excel shape groups Aspose.Cells | worksheet shapes collection iteration | process nested shape groups .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UngroupNestedShapes
{
    // The example loads an Excel workbook, accesses a worksheet, and repeatedly scans its Shapes collection. For each GroupShape it calls Ungroup, looping until no grouped shapes remain, then saves the workbook with all shapes flattened.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            string inputPath = "input.xlsx";
            string outputPath = "output_ungrouped.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Work with the first worksheet (or change index as required)
                Worksheet worksheet = workbook.Worksheets[0];

                // Recursively ungroup all shape groups on this sheet
                UngroupAllShapeGroups(worksheet);

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"All shape groups have been ungrouped. File saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="worksheet">The worksheet to process.</param>
        private static void UngroupAllShapeGroups(Worksheet worksheet)
        {
            if (worksheet == null) return;

            bool groupsRemaining;
            do
            {
                groupsRemaining = false;

                // Work on a copy to avoid collection-modification issues.
                List<Shape> shapesCopy = new List<Shape>();
                foreach (Shape shape in worksheet.Shapes)
                {
                    shapesCopy.Add(shape);
                }

                foreach (Shape shape in shapesCopy)
                {
                    if (shape.IsGroup && shape is GroupShape groupShape)
                    {
                        // Ungroup the current group. After this call, its inner shapes become top‑level shapes.
                        groupShape.Ungroup();
                        groupsRemaining = true; // There may be new groups created from inner shapes.
                    }
                }
            } while (groupsRemaining);
        }
    }
}
