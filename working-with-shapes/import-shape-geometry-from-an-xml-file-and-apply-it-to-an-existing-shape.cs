// Title: Import Shape Geometry from XML into a NotPrimitive AutoShape with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add a NotPrimitive auto shape, read an XML file that contains <Path> elements with MoveTo, LineTo, CurveTo and Close commands, translate the XML into a CustomGeometry object, replace the shape’s existing paths, and save the workbook.
// Keywords: Aspose.Cells | C# | shape geometry import | XML to CustomGeometry | NotPrimitive AutoShape | Excel shape programming | .NET workbook | MoveTo LineTo CurveTo | custom shape paths | programmatic Excel drawing
// Common Searches: Aspose.Cells import shape geometry from XML | C# load custom shape paths into Excel | How to use CustomGeometry with NotPrimitive auto shape | Convert XML path data to Aspose.Cells shape | Add SVG‑like shape to Excel using Aspose.Cells
// Developer Intent: Read an XML definition of a shape’s geometry and assign it to a NotPrimitive auto shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate reusable custom symbols (badges, icons) stored as XML and inject them into multiple reports. | Create complex diagram elements on the fly by importing external geometry definitions. | Standardize corporate branding shapes across workbooks by maintaining a central XML library.
// AI Prompts: Write C# code that parses an XML file with <Path> elements (MoveTo, LineTo, CurveTo, Close) and builds a CustomGeometry for an Aspose.Cells shape. | Explain strategies for validating XML attributes and handling conversion errors when importing shape geometry into Aspose.Cells. | Provide an example XML schema that maps to the Aspose.Cells CustomGeometry API for defining custom shapes.

using System;
using System.IO;
using System.Xml;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeGeometryImportDemo
{
    // Shows how to create a workbook, add a NotPrimitive auto shape, read an XML file that contains <Path> elements with MoveTo, LineTo, CurveTo and Close commands, translate the XML into a CustomGeometry object, replace the shape’s existing paths, and save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a non‑primitive autoshape (initially empty)
                // Parameters: type, upperLeftRow, upperLeftColumn, upperLeftPixelX, upperLeftPixelY, height, width
                Shape shape = sheet.Shapes.AddAutoShape(
                    AutoShapeType.NotPrimitive,
                    2,          // upperLeftRow
                    2,          // upperLeftColumn
                    0,          // upperLeftPixelX
                    0,          // upperLeftPixelY
                    200,        // height (pixels)
                    300);       // width (pixels)

                // Path to the XML file that defines the geometry
                string xmlFilePath = "shapeGeometry.xml";

                // Ensure the XML file exists before loading
                if (!File.Exists(xmlFilePath))
                {
                    Console.WriteLine($"XML file not found: {xmlFilePath}");
                    return;
                }

                // Load and parse the XML
                XmlDocument doc = new XmlDocument();
                doc.Load(xmlFilePath);

                // Cast the shape's geometry to CustomGeometry to access Paths
                CustomGeometry customGeometry = shape.Geometry as CustomGeometry;
                if (customGeometry == null)
                {
                    Console.WriteLine("The shape does not support custom geometry.");
                    return;
                }

                // Clear any existing paths
                customGeometry.Paths.Clear();

                // Iterate over each <Path> element in the XML
                XmlNodeList? pathNodes = doc.SelectNodes("//Geometry/Path");
                if (pathNodes == null) return;

                foreach (XmlNode pathNode in pathNodes)
                {
                    // Create a new path in the collection
                    int pathIndex = customGeometry.Paths.Add();
                    ShapePath path = customGeometry.Paths[pathIndex];

                    // Process child commands (MoveTo, LineTo, CurveTo, Close)
                    foreach (XmlNode cmdNode in pathNode.ChildNodes)
                    {
                        switch (cmdNode.Name)
                        {
                            case "MoveTo":
                                {
                                    float x = float.Parse(cmdNode.Attributes["X"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    float y = float.Parse(cmdNode.Attributes["Y"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    path.MoveTo(x, y);
                                    break;
                                }
                            case "LineTo":
                                {
                                    float x = float.Parse(cmdNode.Attributes["X"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    float y = float.Parse(cmdNode.Attributes["Y"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    path.LineTo(x, y);
                                    break;
                                }
                            case "CurveTo":
                                {
                                    float x1 = float.Parse(cmdNode.Attributes["X1"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    float y1 = float.Parse(cmdNode.Attributes["Y1"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    float x2 = float.Parse(cmdNode.Attributes["X2"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    float y2 = float.Parse(cmdNode.Attributes["Y2"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    float x3 = float.Parse(cmdNode.Attributes["X3"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    float y3 = float.Parse(cmdNode.Attributes["Y3"]?.Value ?? "0", CultureInfo.InvariantCulture);
                                    // Use CubicBezierTo for custom geometry curves
                                    path.CubicBezierTo(x1, y1, x2, y2, x3, y3);
                                    break;
                                }
                            case "Close":
                                {
                                    path.Close();
                                    break;
                                }
                        }
                    }
                }

                // Save the workbook with the updated shape geometry
                string outputPath = "ShapeWithImportedGeometry.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
