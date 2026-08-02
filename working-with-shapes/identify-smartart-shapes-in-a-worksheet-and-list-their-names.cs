// Title: List SmartArt Shape Names in an Excel Worksheet with Aspose.Cells (C#)
// Description: Loads an Excel file, iterates over all worksheet shapes, detects SmartArt objects via the IsSmartArt property, prints each SmartArt shape's Name, and saves the workbook to satisfy the Aspose.Cells lifecycle rule.
// Keywords: Aspose.Cells SmartArt enumeration | C# list SmartArt shapes | Excel SmartArt detection Aspose.Cells | iterate worksheet shapes C# | retrieve SmartArt name Aspose.Cells | Aspose.Cells shape.IsSmartArt | SmartArt shape naming Excel
// Common Searches: How to get SmartArt shape names with Aspose.Cells for .NET | C# code to list SmartArt objects in an Excel worksheet | Aspose.Cells iterate shapes and find SmartArt | Identify SmartArt diagrams in a workbook using C# | Extract SmartArt names from Excel using Aspose.Cells
// Developer Intent: Find every SmartArt object on a worksheet and output its Name property.
// Use Cases: Create an inventory of SmartArt diagrams in a template for documentation. | Validate presence of required SmartArt before running further processing. | Enforce naming conventions by renaming SmartArt shapes after identification. | Generate a report of all SmartArt elements for auditing purposes.
// AI Prompts: Generate C# code that collects all SmartArt shape names from a worksheet into a List<string> using Aspose.Cells. | Show how to filter worksheet shapes by IsSmartArt and export the names to a CSV file. | Provide an example that renames each SmartArt shape after enumeration with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtDemo
{
    // Loads an Excel file, iterates over all worksheet shapes, detects SmartArt objects via the IsSmartArt property, prints each SmartArt shape's Name, and saves the workbook to satisfy the Aspose.Cells lifecycle rule.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
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

            // Save the workbook (even if unchanged) to satisfy lifecycle rule
            workbook.Save("output.xlsx");
        }
    }
}
