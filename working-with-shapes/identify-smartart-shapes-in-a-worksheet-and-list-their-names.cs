// Title: List SmartArt Shape Names in an Excel Worksheet with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, accesses a worksheet, iterates through its Shapes collection, checks the IsSmartArt flag, writes each SmartArt shape's Name to the console, and saves the file.
// Keywords: Aspose.Cells | C# SmartArt shapes | list SmartArt names | IsSmartArt property | Excel worksheet shapes | retrieve SmartArt layout | .NET Excel automation
// Common Searches: How to list SmartArt objects in an Excel file using Aspose.Cells C# | Retrieve names of SmartArt diagrams from a worksheet with Aspose.Cells | Iterate worksheet shapes to find SmartArt in .NET | Get SmartArt shape names from a workbook programmatically
// Developer Intent: Find every SmartArt object in a worksheet and output its name.
// Use Cases: Create an inventory of SmartArt diagrams in a template workbook | Verify required SmartArt exists before generating reports | Log SmartArt identifiers for auditing Excel assets
// AI Prompts: Generate C# code that extracts all SmartArt shape names from every worksheet in a workbook and returns them as a List<string> using Aspose.Cells. | Show how to filter worksheet shapes to SmartArt only and write each name to a CSV file with Aspose.Cells. | Explain how to also obtain the SmartArt layout type and node count for each identified shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtDemo
{
    // Loads a workbook, accesses a worksheet, iterates through its Shapes collection, checks the IsSmartArt flag, writes each SmartArt shape's Name to the console, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or modify as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Check if the shape is a SmartArt object
                if (shape.IsSmartArt)
                {
                    // Output the name of the SmartArt shape
                    Console.WriteLine($"SmartArt Shape Name: {shape.Name}");
                }
            }

            // Save the workbook (no modifications made, but required by lifecycle rules)
            workbook.Save("output.xlsx");
        }
    }
}
