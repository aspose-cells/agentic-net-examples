// Title: Export all worksheet shape definitions to an XML file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, loops through every worksheet, extracts each shape's Name and Type, and writes a custom XML file containing those definitions. | Create a C# method that gathers shape metadata from a Workbook object and saves it as a version‑control‑friendly XML document using Aspose.Cells APIs.
// Common Searches: how to extract shape metadata from an Excel workbook and save as XML using Aspose.Cells C# | Aspose.Cells example for exporting worksheet shapes to an external XML file | C# code to list all shapes in each sheet and generate XML for version control | save Aspose.Cells shape definitions to XML for change tracking
// Tags: aspocells shape xml export | c# extract worksheet shapes aspocells | serialize shape definitions to xml aspocells | version control friendly shape metadata aspocells | iterate workbook shapes c# aspocells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Text;

// Loads an Excel workbook, iterates each worksheet's Shapes collection, builds a simple XML snippet with each shape's Name and Type, and writes the combined XML to ShapesDefinitions.xml for external analysis and version‑control tracking.
class ExportShapesXml
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect XML definitions of all shapes
            StringBuilder xmlBuilder = new StringBuilder();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Build a simple XML representation of the shape
                    string shapeXml = $"<Shape Name=\"{shape.Name}\" Type=\"{shape.Type}\" />";

                    // Add a comment to identify the shape's location
                    xmlBuilder.AppendLine($"<!-- Worksheet: {sheet.Name}, Shape: {shape.Name}, Type: {shape.Type} -->");
                    xmlBuilder.AppendLine(shapeXml);
                    xmlBuilder.AppendLine(); // separate entries
                }
            }

            // Write the collected XML to an external file
            string outputPath = "ShapesDefinitions.xml";
            try
            {
                File.WriteAllText(outputPath, xmlBuilder.ToString());
                Console.WriteLine($"Shape definitions saved to {outputPath}");
            }
            catch (Exception writeEx)
            {
                Console.WriteLine($"Failed to write output file: {writeEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
