// Title: Export worksheet shape geometry (position, size, type) to an XML file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates over every shape on the first worksheet, and writes each shape's name, ID, runtime type, left, top, width, and height into an XML document. | Create a method in a .NET console app that takes a workbook path and an XML output path, ensures the output folder exists, extracts shape geometry from the worksheet using Aspose.Cells, and saves the data as a structured XML file with error handling for individual shapes. | Generate a reusable utility function that serializes non‑primitive shape properties from an Aspose.Cells worksheet into a custom XML schema, including attributes for position and dimensions.
// Common Searches: how to extract Excel shape coordinates with Aspose.Cells C# | Aspose.Cells example for saving shape dimensions to XML | C# code to enumerate worksheet shapes and export their geometry | export non‑primitive shape properties from an .xlsx file to XML using Aspose.Cells | serialize Aspose.Cells shape data to custom XML format
// Tags: Aspose.Cells export shape geometry to XML | enumerate worksheet shapes C# | extract shape position and size Aspose.Cells | serialize Excel shape data as XML | handle shape processing errors Aspose.Cells

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through all shapes on the first worksheet, captures each shape's name, ID, runtime type, and geometric attributes (left, top, width, height), and writes this information into a structured XML file. It creates the output directory if needed and includes per‑shape error handling.
class ExportShapeGeometry
{
    static void Main()
    {
        string workbookPath = "input.xlsx";
        string xmlOutputPath = "shapesGeometry.xml";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Root XML element for all shapes
            XElement root = new XElement("Shapes");

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Basic shape information
                    XElement shapeElement = new XElement("Shape",
                        new XAttribute("Name", shape.Name),
                        new XAttribute("Id", shape.Id),
                        // Use the runtime type name as a fallback for shape type
                        new XAttribute("Type", shape.GetType().Name));

                    // Export position and size as a simple geometry representation
                    XElement geometryElement = new XElement("Geometry",
                        new XAttribute("Left", shape.Left),
                        new XAttribute("Top", shape.Top),
                        new XAttribute("Width", shape.Width),
                        new XAttribute("Height", shape.Height));

                    shapeElement.Add(geometryElement);
                    root.Add(shapeElement);
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Failed to process shape '{shape.Name}': {exShape.Message}");
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(xmlOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the XML document
            XDocument doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
            doc.Save(xmlOutputPath);
            Console.WriteLine($"Shape geometry exported to {xmlOutputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
