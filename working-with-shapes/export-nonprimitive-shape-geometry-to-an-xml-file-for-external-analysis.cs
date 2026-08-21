// Title: Export Non‑Primitive Auto Shape Geometry to XML with Aspose.Cells for .NET (C#)
// Description: The sample creates a new workbook, inserts a non‑primitive auto shape, reads its CustomGeometry, builds an XML document with the shape’s ID, name and path placeholders, and writes the file (NonPrimitiveShapeGeometry.xml) for external processing.
// Keywords: Aspose.Cells | C# | .NET | Excel shape export | non‑primitive auto shape | CustomGeometry | shape geometry XML | shape ID | shape name | XDocument | XML file generation | shape paths extraction
// Common Searches: how to export shape geometry to xml using aspose.cells | c# export non primitive auto shape custom geometry | aspose.cells save shape paths as xml | extract custom geometry from excel shape .net | xml representation of excel shape geometry
// Developer Intent: Generate an XML file that describes the geometry of a non‑primitive auto shape created with Aspose.Cells.
// Use Cases: Provide shape geometry to third‑party diagram or analytics tools. | Archive Excel shape definitions for version‑controlled documentation. | Enable automated inspection of shape paths by parsing the exported XML.
// AI Prompts: Show C# code to retrieve actual path commands from a CustomGeometry object in Aspose.Cells. | Demonstrate how to add fill color and line style information to the exported shape XML. | Explain how to iterate over ShapePath segments and create a detailed XML schema for shape geometry.

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The sample creates a new workbook, inserts a non‑primitive auto shape, reads its CustomGeometry, builds an XML document with the shape’s ID, name and path placeholders, and writes the file (NonPrimitiveShapeGeometry.xml) for external processing.
    public class ExportNonPrimitiveShapeGeometry
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a non‑primitive auto shape to the worksheet
                Shape shape = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.NotPrimitive, 0, 0, 0, 0, 200, 150);

                // Get custom geometry of the shape
                CustomGeometry customGeometry = shape.Geometry as CustomGeometry;
                if (customGeometry == null)
                {
                    Console.WriteLine("The shape does not have custom geometry.");
                    return;
                }

                // Build XML describing the geometry
                XDocument xmlDoc = new XDocument(
                    new XElement("ShapeGeometry",
                        new XAttribute("ShapeId", shape.Id),
                        new XAttribute("ShapeName", shape.Name ?? string.Empty),
                        new XElement("Paths",
                            new Func<XElement>(() =>
                            {
                                XElement pathsElement = new XElement("Paths");
                                for (int i = 0; i < customGeometry.Paths.Count; i++)
                                {
                                    ShapePath path = customGeometry.Paths[i];
                                    pathsElement.Add(
                                        new XElement("Path",
                                            new XAttribute("Index", i),
                                            new XElement("Data", "Path commands not directly accessible via API")
                                        )
                                    );
                                }
                                return pathsElement;
                            })()
                        )
                    )
                );

                // Define output file path
                string outputPath = "NonPrimitiveShapeGeometry.xml";

                // Write the XML to file
                string xmlContent = (xmlDoc.Declaration?.ToString() ?? string.Empty) + Environment.NewLine + xmlDoc.ToString();
                File.WriteAllText(outputPath, xmlContent);

                Console.WriteLine($"Shape geometry exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportNonPrimitiveShapeGeometry.Run();
        }
    }
}
