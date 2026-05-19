using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class CompareShapeGeometry
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 50);

        // Capture geometry adjustment values before XML import
        Geometry geometryBefore = shape.Geometry;
        List<double> adjustValuesBefore = new List<double>();
        foreach (ShapeGuide guide in geometryBefore.ShapeAdjustValues)
        {
            adjustValuesBefore.Add(guide.Value);
        }

        // Prepare a simple XML document in a memory stream
        string xmlContent = "<root><data>123</data></root>";
        MemoryStream xmlStream = new MemoryStream();
        using (StreamWriter writer = new StreamWriter(xmlStream, System.Text.Encoding.UTF8, 1024, true))
        {
            writer.Write(xmlContent);
            writer.Flush();
            xmlStream.Position = 0;
        }

        // Import the XML data into the workbook starting at cell A1 of Sheet1
        workbook.ImportXml(xmlStream, "Sheet1", 0, 0);

        // Capture geometry adjustment values after XML import
        Geometry geometryAfter = shape.Geometry;
        List<double> adjustValuesAfter = new List<double>();
        foreach (ShapeGuide guide in geometryAfter.ShapeAdjustValues)
        {
            adjustValuesAfter.Add(guide.Value);
        }

        // Compare the two sets of adjustment values for consistency
        bool geometryIsConsistent = true;
        if (adjustValuesBefore.Count != adjustValuesAfter.Count)
        {
            geometryIsConsistent = false;
        }
        else
        {
            for (int i = 0; i < adjustValuesBefore.Count; i++)
            {
                if (Math.Abs(adjustValuesBefore[i] - adjustValuesAfter[i]) > 1e-6)
                {
                    geometryIsConsistent = false;
                    break;
                }
            }
        }

        Console.WriteLine("Geometry unchanged after XML import: " + geometryIsConsistent);

        // Save the workbook to verify the final state
        workbook.Save("CompareGeometry.xlsx");
    }
}