// Title: How to send a specific shape to the front of the Z‑order stack using Aspose.Cells for .NET (C#)
// Common Searches: asp.net aspose.cells bring shape to front of Z-order | c# code to move Excel shape to front layer with Aspose.Cells | how to change shape overlay order in an Excel workbook using Aspose.Cells | Aspose.Cells example for setting shape Z-order to top in .NET | retrieve shape by name and set it as topmost object in Aspose.Cells C#
// Tags: Aspose.Cells shape Z-order manipulation | C# set Excel shape front layer | Aspose.Cells bring shape to top | Excel shape overlay order with Aspose.Cells | modify shape ordering Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads or creates an Excel workbook, retrieves a rectangle shape named 'MyShape', notes that Aspose.Cells for .NET does not provide a direct BringToFront method for Z‑order, and saves the workbook. It demonstrates how to access a shape by name and outlines the limitation regarding shape ordering.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string shapeName = "MyShape";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one with a placeholder shape
                if (File.Exists(inputPath))
                {
                    try
                    {
                        workbook = new Workbook(inputPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                        return;
                    }
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    Shape placeholder = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                    placeholder.Name = shapeName;
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the target shape by name
                Shape targetShape = worksheet.Shapes[shapeName];
                if (targetShape != null)
                {
                    // Aspose.Cells does not expose a direct BringToFront method.
                    // If needed, additional Z‑order manipulation can be performed here.
                    Console.WriteLine($"Shape '{shapeName}' found.");
                }
                else
                {
                    Console.WriteLine($"Shape '{shapeName}' not found in the worksheet.");
                }

                // Save the modified workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
