// Title: C# Script to Validate and Resize Oversized Shapes in an Excel Workbook with Aspose.Cells
// Description: Loads an Excel file, iterates through each worksheet’s ShapeCollection, identifies shapes whose Width or Height exceed user‑defined pixel limits, logs the violations, resizes the shapes to the allowed dimensions, and saves the updated workbook.
// Keywords: Aspose.Cells shape validation | C# resize Excel shapes | limit shape width height | shape dimension check Aspose | automated shape resizing .NET | Excel workbook shape size enforcement | shape collection iteration Aspose.Cells
// Common Searches: how to enforce maximum shape size in Excel using Aspose.Cells | C# script to resize shapes that are too large in a workbook | validate shape dimensions across all worksheets with Aspose.Cells | restrict shape width and height in Excel programmatically | detect and adjust oversized images in Excel files .NET
// Developer Intent: Detect any shape that exceeds the specified width or height and automatically resize it to comply with the defined limits.
// Use Cases: Ensure template consistency by preventing shapes from breaking layout boundaries. | Automate compliance with corporate branding rules that cap image dimensions. | Pre‑process user‑generated workbooks to avoid printing or rendering issues caused by large shapes. | Integrate shape size checks into a CI pipeline for Excel report generation.
// AI Prompts: Create a version of ShapeValidator that writes adjustment details to a CSV log file. | Modify the script to keep the original aspect ratio when only one dimension exceeds the limit. | Generate unit tests that verify shapes are correctly resized when they surpass the maximum width or height. | Adapt the code to accept size limits in points instead of pixels. | Provide a PowerShell wrapper that runs the ShapeValidator on multiple files in a folder.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeValidationDemo
{
    // Loads an Excel file, iterates through each worksheet’s ShapeCollection, identifies shapes whose Width or Height exceed user‑defined pixel limits, logs the violations, resizes the shapes to the allowed dimensions, and saves the updated workbook.
    public class ShapeValidator
    {
        /// <param name="inputFile">Path to the source Excel file.</param>
        /// <param name="outputFile">Path where the validated workbook will be saved.</param>
        /// <param name="maxWidth">Maximum allowed width in pixels.</param>
        /// <param name="maxHeight">Maximum allowed height in pixels.</param>
        public static void Run(string inputFile, string outputFile, int maxWidth, int maxHeight)
        {
            // Load the workbook using the standard constructor (lifecycle rule)
            Workbook workbook = new Workbook(inputFile, new LoadOptions());

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Examine each shape individually
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Check width and height against the thresholds
                    bool widthExceeded = shape.Width > maxWidth;
                    bool heightExceeded = shape.Height > maxHeight;

                    if (widthExceeded || heightExceeded)
                    {
                        // Output diagnostic information
                        Console.WriteLine($"Worksheet '{sheet.Name}', Shape {i} ('{shape.Name}') exceeds limits:");
                        Console.WriteLine($"  Current Size -> Width: {shape.Width}px, Height: {shape.Height}px");
                        Console.WriteLine($"  Limits        -> Max Width: {maxWidth}px, Max Height: {maxHeight}px");

                        // Resize the shape to fit within the allowed dimensions
                        if (widthExceeded)
                        {
                            shape.Width = maxWidth;
                        }

                        if (heightExceeded)
                        {
                            shape.Height = maxHeight;
                        }

                        Console.WriteLine($"  Adjusted Size -> Width: {shape.Width}px, Height: {shape.Height}px");
                    }
                }
            }

            // Save the validated workbook using the standard Save method (lifecycle rule)
            workbook.Save(outputFile);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourcePath = "InputWorkbook.xlsx";
            string destinationPath = "ValidatedWorkbook.xlsx";
            int allowedWidth = 300;   // pixels
            int allowedHeight = 200;  // pixels

            ShapeValidator.Run(sourcePath, destinationPath, allowedWidth, allowedHeight);
        }
    }
}
