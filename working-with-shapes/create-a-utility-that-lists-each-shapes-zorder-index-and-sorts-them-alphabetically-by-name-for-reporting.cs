using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsUtilities
{
    /// <summary>
    /// Utility that reports each shape's Z‑order index and lists them alphabetically by name.
    /// </summary>
    public static class ShapeZOrderReporter
    {
        /// <summary>
        /// Loads a workbook, gathers shape information, sorts it by name, prints the report,
        /// and saves the workbook (optional).
        /// </summary>
        /// <param name="inputPath">Path to the Excel file to load. If null or empty, a new workbook is created.</param>
        /// <param name="outputPath">Path to save the workbook after processing. If null or empty, the workbook is not saved.</param>
        public static void Run(string inputPath, string outputPath)
        {
            // Load or create workbook
            Workbook workbook;
            if (!string.IsNullOrEmpty(inputPath))
            {
                // Load existing workbook
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook and add sample shapes for demonstration
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];

                // Add three shapes with default names
                ws.Shapes.AddRectangle(2, 0, 2, 0, 80, 120);   // "Rectangle 1"
                ws.Shapes.AddOval(5, 0, 5, 0, 80, 120);       // "Oval 1"
                ws.Shapes.AddLine(8, 0, 8, 0, 80, 120);       // "Line 1"

                // Optionally rename shapes to illustrate alphabetical sorting
                ws.Shapes[0].Name = "AlphaShape";
                ws.Shapes[1].Name = "BetaShape";
                ws.Shapes[2].Name = "GammaShape";

                // Adjust Z‑order positions for demonstration
                ws.Shapes[0].ZOrderPosition = 2; // front
                ws.Shapes[1].ZOrderPosition = 0; // back
                ws.Shapes[2].ZOrderPosition = 1; // middle
            }

            // Work with the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Collect name and Z‑order for each shape
            var shapeInfoList = new List<(string Name, int ZOrder)>();
            foreach (Shape shape in shapes)
            {
                // Some shapes may have empty names; use a placeholder if needed
                string name = string.IsNullOrEmpty(shape.Name) ? $"Unnamed_{shape.Id}" : shape.Name;
                shapeInfoList.Add((name, shape.ZOrderPosition));
            }

            // Sort the list alphabetically by shape name
            shapeInfoList.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));

            // Output the report
            Console.WriteLine("Shape Z‑Order Report (sorted by name):");
            Console.WriteLine("--------------------------------------");
            foreach (var info in shapeInfoList)
            {
                Console.WriteLine($"Name: {info.Name}, Z‑Order Position: {info.ZOrder}");
            }

            // Save the workbook if an output path is provided
            if (!string.IsNullOrEmpty(outputPath))
            {
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to: {outputPath}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            // Provide paths as needed; empty strings will trigger creation of a sample workbook.
            string inputFile = "";   // e.g., "input.xlsx"
            string outputFile = "ShapeReportDemo.xlsx";

            ShapeZOrderReporter.Run(inputFile, outputFile);
        }
    }
}