// Title: Export non‑primitive shape geometry from an Excel workbook to an XML file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates over every shape on each worksheet, and writes each shape's name, type, top, left, height, and width into a structured XML document. | Create a method that extracts shape properties from a Worksheet.Shapes collection and returns an XDocument containing <Shape> elements with geometry attributes, while handling exceptions for unsupported shapes. | Write a console application that validates the input Excel file, exports all shape geometry to a file named shapes_geometry.xml, and logs warnings for any shapes that cannot be processed.
// Common Searches: how to extract shape positions from an Excel file using Aspose.Cells in C# | Aspose.Cells C# export shape dimensions to XML | C# code to iterate worksheet shapes and save their geometry as XML | save Excel shape metadata to XML with Aspose.Cells library | export non‑primitive shapes from .xlsx to XML using Aspose.Cells
// Tags: Aspose.Cells export shape geometry to XML | C# iterate worksheet shapes Aspose.Cells | shape attribute extraction Excel .xlsx | XML serialization of Excel shape properties | error handling for unsupported shapes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Xml.Linq;

// The example loads an Excel workbook with Aspose.Cells, checks that the file exists, iterates through each worksheet and its Shapes collection, captures each shape's name, type, top, left, height, and width, creates <Shape> elements with these attributes under a root <Shapes> node, handles per‑shape exceptions, and saves the resulting XML to shapes_geometry.xml.
class ExportShapeGeometry
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "shapes_geometry.xml";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Root XML element for all shapes
            XElement root = new XElement("Shapes");

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Create XML element for the shape with basic attributes
                        XElement shapeElement = new XElement("Shape",
                            new XAttribute("Name", shape.Name ?? string.Empty),
                            new XAttribute("Type", shape.Type.ToString()),
                            new XAttribute("Top", shape.Top),
                            new XAttribute("Left", shape.Left),
                            new XAttribute("Height", shape.Height),
                            new XAttribute("Width", shape.Width));

                        // Add the shape element to the root
                        root.Add(shapeElement);
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Warning: Could not process shape '{shape.Name}'. Details: {shapeEx.Message}");
                    }
                }
            }

            // Save the XML document
            XDocument doc = new XDocument(root);
            doc.Save(outputFile);
            Console.WriteLine($"Shape information exported successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
