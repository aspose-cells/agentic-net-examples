// Title: Import Shape Geometry from XML into an Aspose.Cells AutoShape (C#)
// Description: Creates a workbook, adds a non‑primitive AutoShape, casts its Geometry to CustomGeometry, reads an XML file with <Path> elements, translates MoveTo, LineTo, CurveTo and Close commands into ShapePath instructions, and saves the workbook with the new geometry.
// Keywords: Aspose.Cells | C# | Import shape geometry | CustomGeometry | AutoShape XML | ShapePath MoveTo | ShapePath LineTo | CubicBezierTo | Excel vector shape
// Common Searches: Aspose.Cells import shape geometry from XML | C# add custom geometry to AutoShape | How to parse XML shape paths in Aspose.Cells | Load XML path commands into Excel shape | Custom shape from XML using Aspose.Cells
// Developer Intent: Read vector path data defined in an XML file and apply it to a custom‑geometry AutoShape in an Excel workbook using Aspose.Cells for C#.
// Use Cases: Convert external SVG‑like XML definitions into precise Excel shapes for reporting dashboards. | Replace placeholder symbols in templates with geometry supplied by configuration files. | Build a migration tool that reproduces proprietary vector graphics in Excel workbooks.
// AI Prompts: Generate a reusable C# method that accepts an XML file path and a CustomGeometry object, parses <Path> elements, and populates MoveTo, LineTo, CubicBezierTo, and Close commands. | Create code to translate SVG path data (M, L, C, Z) into Aspose.Cells ShapePath calls for bulk shape import. | Suggest comprehensive error‑handling for missing or malformed XML attributes when importing shape geometry into Aspose.Cells.

using System;
using System.Globalization;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGeometryImport
{
    // Creates a workbook, adds a non‑primitive AutoShape, casts its Geometry to CustomGeometry, reads an XML file with <Path> elements, translates MoveTo, LineTo, CurveTo and Close commands into ShapePath instructions, and saves the workbook with the new geometry.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a non‑primitive auto shape (custom geometry placeholder)
                // Parameters: shape type, upper left row/col, upper left offset, width, height
                Shape shape = sheet.Shapes.AddAutoShape(
                    AutoShapeType.NotPrimitive, // custom geometry shape
                    1, 1,   // upper left row, column
                    0, 0,   // offset within the cell
                    200, 200); // width, height in pixels

                // Cast the shape's geometry to CustomGeometry to access the Paths collection
                CustomGeometry? customGeometry = shape.Geometry as CustomGeometry;
                if (customGeometry == null)
                {
                    Console.WriteLine("The shape does not support custom geometry.");
                    return;
                }

                // Load the XML that defines the geometry.
                string xmlPath = "shapeGeometry.xml";
                if (!File.Exists(xmlPath))
                {
                    Console.WriteLine($"File not found: {xmlPath}");
                    return;
                }

                XDocument doc = XDocument.Load(xmlPath);

                // Iterate over each <Path> element in the XML
                foreach (XElement pathElement in doc.Root?.Elements("Path") ?? Array.Empty<XElement>())
                {
                    // Add a new path to the shape's geometry
                    int pathIndex = customGeometry.Paths.Add();
                    ShapePath path = customGeometry.Paths[pathIndex];

                    // Process child commands of the path (MoveTo, LineTo, CurveTo, Close, etc.)
                    foreach (XElement cmd in pathElement.Elements())
                    {
                        switch (cmd.Name.LocalName)
                        {
                            case "MoveTo":
                                // MoveTo X="value" Y="value"
                                float moveX = float.Parse(cmd.Attribute("X")?.Value ?? "0", CultureInfo.InvariantCulture);
                                float moveY = float.Parse(cmd.Attribute("Y")?.Value ?? "0", CultureInfo.InvariantCulture);
                                path.MoveTo(moveX, moveY);
                                break;

                            case "LineTo":
                                // LineTo X="value" Y="value"
                                float lineX = float.Parse(cmd.Attribute("X")?.Value ?? "0", CultureInfo.InvariantCulture);
                                float lineY = float.Parse(cmd.Attribute("Y")?.Value ?? "0", CultureInfo.InvariantCulture);
                                path.LineTo(lineX, lineY);
                                break;

                            case "CurveTo":
                                // CurveTo X1=".." Y1=".." X2=".." Y2=".." X3=".." Y3=".."
                                float x1 = float.Parse(cmd.Attribute("X1")?.Value ?? "0", CultureInfo.InvariantCulture);
                                float y1 = float.Parse(cmd.Attribute("Y1")?.Value ?? "0", CultureInfo.InvariantCulture);
                                float x2 = float.Parse(cmd.Attribute("X2")?.Value ?? "0", CultureInfo.InvariantCulture);
                                float y2 = float.Parse(cmd.Attribute("Y2")?.Value ?? "0", CultureInfo.InvariantCulture);
                                float x3 = float.Parse(cmd.Attribute("X3")?.Value ?? "0", CultureInfo.InvariantCulture);
                                float y3 = float.Parse(cmd.Attribute("Y3")?.Value ?? "0", CultureInfo.InvariantCulture);
                                // Use CubicBezierTo if CurveTo is unavailable
                                path.CubicBezierTo(x1, y1, x2, y2, x3, y3);
                                break;

                            case "Close":
                                // Close the current path
                                path.Close();
                                break;

                            default:
                                Console.WriteLine($"Unsupported command: {cmd.Name}");
                                break;
                        }
                    }
                }

                // Save the workbook with the updated shape geometry
                string outputPath = "ShapeWithImportedGeometry.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Shape geometry imported and workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
