// Title: How to programmatically refresh a Camera shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, finds a shape named "Camera 1" on a worksheet, invokes its Refresh method via dynamic typing, and saves the workbook. | Show an example of safely calling the Refresh method on a CameraShape when the exact shape type is unknown at compile time in Aspose.Cells. | Write a C# snippet that verifies a camera shape exists, refreshes its view, handles the case where Refresh is unavailable, and then persists the changes.
// Common Searches: Aspose.Cells C# refresh camera shape after workbook load | How to call Refresh on Excel CameraShape using Aspose.Cells .NET API | Dynamic invocation of Refresh method on a shape in Aspose.Cells | Programmatically update camera view in an Excel file with Aspose.Cells | C# locate and refresh a named shape in an Excel worksheet using Aspose.Cells
// Tags: camera shape manipulation Aspose.Cells C# | find shape by name Aspose.Cells | dynamic call of shape methods Aspose.Cells | update Excel camera view programmatically | save workbook after shape operation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, searches the first worksheet for a shape called "Camera 1", attempts to invoke its Refresh method using dynamic typing (ignoring the call if unsupported), and then saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string targetShapeName = "Camera 1";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the shape by name
            Shape targetShape = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape.Name.Equals(targetShapeName, StringComparison.OrdinalIgnoreCase))
                {
                    targetShape = shape;
                    break;
                }
            }

            // If the shape is found and it supports Refresh (e.g., CameraShape), invoke it safely
            if (targetShape != null)
            {
                try
                {
                    // Use dynamic to call Refresh without compile‑time type dependency
                    dynamic dynShape = targetShape;
                    dynShape.Refresh();
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // Refresh method not available on this shape type; ignore safely
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while refreshing shape: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Shape named \"{targetShapeName}\" not found.");
            }

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
