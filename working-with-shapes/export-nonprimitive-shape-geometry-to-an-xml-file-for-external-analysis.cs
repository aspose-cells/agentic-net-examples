// Title: Export Non‑Primitive AutoShape Geometry to XML with Aspise.Cells for .NET
// Description: Creates a workbook, adds a NotPrimitive auto shape, extracts its CustomGeometry, and writes the shape ID, type, adjustment values and path indexes to a formatted XML file, then saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | Export shape geometry | Non‑primitive AutoShape | CustomGeometry | XML output | Shape adjustment values | Shape paths | Workbook export | External shape analysis
// Common Searches: Aspose.Cells export shape geometry to XML | How to write NotPrimitive auto shape data to XML in C# | Retrieve custom geometry of a shape using Aspose.Cells | Save shape adjustment values as XML with Aspose.Cells | Export AutoShape paths from Excel workbook
// Developer Intent: Generate an XML representation of a non‑primitive auto shape’s geometry for downstream processing or validation.
// Use Cases: Produce an XML report of shape adjustments for design QA. | Create version‑controlled snapshots of custom shape geometry to detect changes over time. | Supply shape geometry data to external tools that consume XML for rendering or analysis.
// AI Prompts: Write C# code that reads the ShapeGeometry.xml produced by this sample and rebuilds the same NotPrimitive shape in a new workbook using Aspose.Cells. | Suggest a technique to capture detailed path segment commands (MoveTo, LineTo, etc.) from CustomGeometry despite the current API limitation. | Explain how to loop through all shapes on a worksheet and export each shape’s geometry to separate XML files.

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGeometryExport
{
    // Creates a workbook, adds a NotPrimitive auto shape, extracts its CustomGeometry, and writes the shape ID, type, adjustment values and path indexes to a formatted XML file, then saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a non‑primitive auto shape (NotPrimitive) – this shape has custom geometry
                // Parameters: shape type, upper left row, upper left column, top, left, width, height
                Shape shape = sheet.Shapes.AddAutoShape(
                    AutoShapeType.NotPrimitive, // non‑primitive shape
                    0, 0,                      // upper left row, column
                    0, 0,                      // top, left (in pixels)
                    200, 300);                 // width, height (in pixels)

                // Cast the geometry to CustomGeometry to access the Paths collection
                CustomGeometry customGeometry = shape.Geometry as CustomGeometry;
                if (customGeometry == null)
                {
                    Console.WriteLine("The shape does not have custom geometry.");
                    return;
                }

                // Prepare an XML file to store the geometry information
                string xmlPath = "ShapeGeometry.xml";

                // Ensure the directory for the XML file exists (if a directory is specified)
                string xmlDir = Path.GetDirectoryName(Path.GetFullPath(xmlPath));
                if (!string.IsNullOrEmpty(xmlDir) && !Directory.Exists(xmlDir))
                {
                    Directory.CreateDirectory(xmlDir);
                }

                using (XmlWriter writer = XmlWriter.Create(xmlPath, new XmlWriterSettings { Indent = true }))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("ShapeGeometry");

                    // Basic shape information
                    writer.WriteAttributeString("Id", shape.Id.ToString());
                    writer.WriteAttributeString("AutoShapeType", shape.AutoShapeType.ToString());

                    // Write adjustment values (if any)
                    if (shape.Geometry != null && shape.Geometry.ShapeAdjustValues != null && shape.Geometry.ShapeAdjustValues.Count > 0)
                    {
                        writer.WriteStartElement("AdjustValues");
                        for (int i = 0; i < shape.Geometry.ShapeAdjustValues.Count; i++)
                        {
                            var adj = shape.Geometry.ShapeAdjustValues[i];
                            writer.WriteStartElement("Adjust");
                            writer.WriteAttributeString("Index", i.ToString());
                            // Some versions expose a Name property; fall back to index if not available
                            try
                            {
                                var nameProp = adj.GetType().GetProperty("Name");
                                if (nameProp != null)
                                {
                                    string name = nameProp.GetValue(adj)?.ToString() ?? "";
                                    writer.WriteAttributeString("Name", name);
                                }
                            }
                            catch { /* ignore reflection errors */ }

                            writer.WriteAttributeString("Value", adj.Value.ToString());
                            writer.WriteEndElement(); // Adjust
                        }
                        writer.WriteEndElement(); // AdjustValues
                    }

                    // Write path information
                    writer.WriteStartElement("Paths");
                    for (int i = 0; i < customGeometry.Paths.Count; i++)
                    {
                        writer.WriteStartElement("Path");
                        writer.WriteAttributeString("Index", i.ToString());

                        // Note: Aspose.Cells does not expose the segment list directly,
                        // so we record the existence of the path and its index.
                        writer.WriteComment("Path commands (MoveTo, LineTo, etc.) are not directly readable via the API.");

                        writer.WriteEndElement(); // Path
                    }
                    writer.WriteEndElement(); // Paths

                    writer.WriteEndElement(); // ShapeGeometry
                    writer.WriteEndDocument();
                }

                Console.WriteLine($"Shape geometry exported to '{xmlPath}'.");

                // Save the workbook (lifecycle rule: save)
                string workbookPath = "ShapeWithGeometry.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
