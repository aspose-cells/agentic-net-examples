// Title: C# – Log SmartArt Shape Type, Position and Size in an Excel Workbook with Aspose.Cells
// Description: Sample code that opens an Excel file, iterates through all worksheets, identifies SmartArt shapes, and writes each shape’s name, .NET type, left/top coordinates and width/height (pixels) to the console. Shows how to use Aspose.Cells ShapeCollection and SmartArtShape classes.
// Keywords: Aspose.Cells SmartArt | C# SmartArt shape properties | retrieve SmartArt position | SmartArt size pixels | log SmartArt type | Excel shape collection .NET | iterate worksheet shapes | Aspose.Cells get SmartArt dimensions | SmartArt shape logging C# | Excel automation Aspose.Cells
// Common Searches: How to list SmartArt shapes in an Excel file using Aspose.Cells C# | Get SmartArt coordinates and size with Aspose.Cells .NET | Retrieve SmartArt type and name from worksheets in C# | Log SmartArt shape details while iterating worksheet shapes | Aspose.Cells example for SmartArt position and dimensions
// Developer Intent: Extract and log the name, .NET type, left/top position and width/height of every SmartArt shape in an Excel workbook.
// Use Cases: Audit SmartArt objects to ensure correct layout before publishing a workbook. | Create a CSV report of SmartArt locations and dimensions for automated slide conversion. | Validate SmartArt size and placement against corporate design guidelines during batch processing. | Debug layout issues by printing shape metrics during development.
// AI Prompts: Generate C# code that exports SmartArt shape details (name, type, position, size) to a CSV file using Aspose.Cells. | Show how to filter SmartArt shapes by a specific layout type (e.g., Process) while iterating worksheets. | Explain how to change the position or size of a SmartArt shape after logging its current values with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtInfoLogger
{
    // Sample code that opens an Excel file, iterates through all worksheets, identifies SmartArt shapes, and writes each shape’s name, .NET type, left/top coordinates and width/height (pixels) to the console. Shows how to use Aspose.Cells ShapeCollection and SmartArtShape classes.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Loop through each shape
                foreach (Shape shape in shapes)
                {
                    // Check if the shape is a SmartArt shape
                    if (shape.IsSmartArt && shape is SmartArtShape smartArt)
                    {
                        // Log detailed information
                        Console.WriteLine("Worksheet: {0}", sheet.Name);
                        Console.WriteLine("SmartArt Name: {0}", smartArt.Name);
                        Console.WriteLine("Type: {0}", smartArt.GetType().Name);
                        Console.WriteLine("Position - Left: {0} px, Top: {1} px", smartArt.Left, smartArt.Top);
                        Console.WriteLine("Size - Width: {0} px, Height: {1} px", smartArt.Width, smartArt.Height);
                        Console.WriteLine(new string('-', 40));
                    }
                }
            }

            // Optionally save the workbook if any modifications were made
            workbook.Save("output.xlsx");
        }
    }
}
