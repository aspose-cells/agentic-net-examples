// Title: Import shape geometry from an XML file and apply it to a shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# load custom shape geometry from XML and apply to Aspose.Cells shape | Aspose.Cells set shape custom geometry using an XML definition | use reflection to access the CustomGeometry property of a shape in Aspose.Cells | import shape geometry XML into an existing Excel shape with Aspose.Cells .NET | apply XML-defined geometry to the first shape on a worksheet using Aspose.Cells
// Tags: load custom geometry XML Aspose.Cells | set shape CustomGeometry property C# | apply XML-defined shape geometry .NET | modify Excel shape geometry using Aspose.Cells | reflection-based access to CustomGeometry Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGeometryExample
{
    // The example loads an existing workbook, reads a custom shape geometry definition from an XML file, uses reflection to locate the shape's CustomGeometry property and its LoadXml method, applies the XML geometry to the first shape on the first worksheet, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputWorkbookPath = "InputWorkbook.xlsx";
                const string geometryXmlPath = "ShapeGeometry.xml";
                const string outputWorkbookPath = "OutputWorkbook.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputWorkbookPath))
                {
                    Console.WriteLine($"Error: Input workbook not found at '{inputWorkbookPath}'.");
                    return;
                }

                // Verify that the geometry XML file exists
                if (!File.Exists(geometryXmlPath))
                {
                    Console.WriteLine($"Error: Geometry XML file not found at '{geometryXmlPath}'.");
                    return;
                }

                // Load the workbook that contains the target shape
                Workbook workbook = new Workbook(inputWorkbookPath);

                // Access the first worksheet (adjust as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one shape
                if (sheet.Shapes.Count == 0)
                {
                    Console.WriteLine("Error: No shapes found on the first worksheet.");
                    return;
                }

                // Retrieve the shape you want to modify (by index or name)
                // Here we assume the shape is the first shape on the sheet
                Shape shape = sheet.Shapes[0];

                // Read the custom geometry XML from the file
                string geometryXml;
                try
                {
                    geometryXml = File.ReadAllText(geometryXmlPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading geometry XML: {ex.Message}");
                    return;
                }

                // Apply the custom geometry to the shape.
                // Use reflection to stay compatible with versions that may not expose CustomGeometry directly.
                try
                {
                    var customGeometryProp = shape.GetType().GetProperty("CustomGeometry");
                    if (customGeometryProp != null)
                    {
                        var customGeometryObj = customGeometryProp.GetValue(shape);
                        var loadXmlMethod = customGeometryObj?.GetType().GetMethod("LoadXml");
                        loadXmlMethod?.Invoke(customGeometryObj, new object[] { geometryXml });
                    }
                    else
                    {
                        Console.WriteLine("Custom geometry feature is not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error applying custom geometry: {ex.Message}");
                    return;
                }

                // Save the workbook with the updated shape geometry
                workbook.Save(outputWorkbookPath);
                Console.WriteLine($"Workbook saved successfully to '{outputWorkbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
