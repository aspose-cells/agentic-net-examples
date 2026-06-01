using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ExportShapeGeometry
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a non‑primitive auto shape (custom geometry)
            // Parameters: AutoShapeType.NotPrimitive, upper left row, column, upper left pixel offset X/Y, width, height
            Shape shape = worksheet.Shapes.AddAutoShape(AutoShapeType.NotPrimitive, 0, 0, 0, 0, 200, 200);

            // Access the custom geometry of the shape
            CustomGeometry customGeometry = shape.Geometry as CustomGeometry;
            if (customGeometry == null)
            {
                Console.WriteLine("The shape does not contain custom geometry.");
                return;
            }

            // Build a simple rectangular path
            int pathIndex = customGeometry.Paths.Add(); // Create a new path
            ShapePath path = customGeometry.Paths[pathIndex];
            path.MoveTo(0, 0);
            path.LineTo(20000, 0);
            path.LineTo(20000, 20000);
            path.LineTo(0, 20000);
            path.Close();

            // Export geometry information to an XML file
            string xmlFilePath = "ShapeGeometry.xml";
            using (XmlWriter writer = XmlWriter.Create(xmlFilePath, new XmlWriterSettings { Indent = true }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("ShapeGeometry");
                writer.WriteAttributeString("ShapeId", shape.Id.ToString());
                writer.WriteAttributeString("AutoShapeType", shape.AutoShapeType.ToString());

                // Write adjustment values (if any)
                writer.WriteStartElement("AdjustmentValues");
                foreach (var adj in shape.Geometry.ShapeAdjustValues)
                {
                    writer.WriteStartElement("Adjustment");
                    // Some versions expose a Name property; if not, only write the value.
                    writer.WriteAttributeString("Value", adj.Value.ToString());
                    writer.WriteEndElement(); // Adjustment
                }
                writer.WriteEndElement(); // AdjustmentValues

                // Write each path and its commands
                writer.WriteStartElement("Paths");
                for (int i = 0; i < customGeometry.Paths.Count; i++)
                {
                    ShapePath sp = customGeometry.Paths[i];
                    writer.WriteStartElement("Path");
                    writer.WriteAttributeString("Index", i.ToString());

                    // Export raw path data as a string
                    writer.WriteElementString("PathData", sp.ToString());

                    writer.WriteEndElement(); // Path
                }
                writer.WriteEndElement(); // Paths

                writer.WriteEndElement(); // ShapeGeometry
                writer.WriteEndDocument();
            }

            Console.WriteLine($"Shape geometry exported to '{xmlFilePath}'.");

            // Save the workbook
            string workbookPath = "ExportShapeGeometryDemo.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to '{workbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}