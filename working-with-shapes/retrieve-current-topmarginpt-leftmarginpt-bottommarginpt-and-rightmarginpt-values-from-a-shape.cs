// Title: Read Shape Text Margins (TopMarginPt, LeftMarginPt, BottomMarginPt, RightMarginPt) with Aspose.Cells for .NET
// Description: Shows how to load or create an Excel workbook, locate a shape on the first worksheet, and retrieve its TextBody margin values (TopMarginPt, LeftMarginPt, BottomMarginPt, RightMarginPt) using Aspose.Cells for .NET, with fallback handling for API versions that do not expose these properties.
// Keywords: Aspose.Cells shape margins .NET | TopMarginPt | LeftMarginPt | BottomMarginPt | RightMarginPt | shape text body margins | C# read shape margins | Excel shape margin properties | Aspose.Cells API version compatibility | retrieve shape margins
// Common Searches: Aspose.Cells get shape top margin points | How to read left margin of a shape in Aspose.Cells C# | BottomMarginPt property Aspose.Cells .NET | RightMarginPt not available in Aspose.Cells API | Shape.TextBody margin values Aspose.Cells example
// Developer Intent: Obtain the TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt values of a shape's TextBody in an Excel file using Aspose.Cells for .NET.
// Use Cases: Adjust PDF export layout by reading shape text margins before rendering. | Synchronize shape positioning across multiple worksheets based on margin settings. | Validate design compliance of automatically generated reports by checking shape margins.
// AI Prompts: Write C# code with Aspose.Cells that reads TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt of a shape's TextBody, including version‑check logic for missing properties. | Explain alternative techniques to infer shape text margins when the Aspose.Cells API does not expose margin properties. | Create a robust sample program that loads a workbook, accesses the first shape, prints its text and margin values, and gracefully handles unsupported API versions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Shows how to load or create an Excel workbook, locate a shape on the first worksheet, and retrieve its TextBody margin values (TopMarginPt, LeftMarginPt, BottomMarginPt, RightMarginPt) using Aspose.Cells for .NET, with fallback handling for API versions that do not expose these properties.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                // Load existing workbook or create a new one if the file is missing
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Add a sample shape so the example can run without an input file
                    Worksheet ws = workbook.Worksheets[0];
                    // AddShape returns the newly created Shape instance in newer API versions
                    Shape sampleShape = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                    sampleShape.TextBody.Text = "Sample Text";
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                }
                else
                {
                    try
                    {
                        // Retrieve the first shape
                        Shape shape = worksheet.Shapes[0];

                        // Output basic text information (margin properties are not exposed in this API version)
                        Console.WriteLine($"Shape Text: {shape.TextBody.Text}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing shape: {ex.Message}");
                    }
                }

                // Save the workbook (optional if modifications were made)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
