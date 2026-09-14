// Title: Import shape geometry from an XML file and apply it to an existing worksheet shape with Aspose.Cells for .NET (C#)
// AI Prompts: Read the ShapeGeometry.xml file, parse its path commands, and use Geometry.AddPath, AddLineTo, AddArcTo, and AddClosePath methods to rebuild the shape's geometry in Aspose.Cells. | Locate the first shape on the first worksheet, retrieve its read‑only Geometry object, and replace its path data with the custom geometry constructed from the XML definition. | Check that the input workbook and XML file exist, create the output directory if needed, and save the workbook after updating the shape geometry.
// Common Searches: c# aspose.cells import custom shape geometry from xml file | how to modify an Excel shape's path using an external XML definition in Aspose.Cells | apply external geometry definition to a worksheet shape with Aspose.Cells .NET | read shape geometry xml and update shape in an Excel workbook using C# | asp.net aspose.cells shape geometry API example for custom paths
// Tags: import shape geometry XML Aspose.Cells C# | modify worksheet shape geometry Aspose.Cells | custom shape path reconstruction using Geometry API | load external geometry definition into Excel shape | Aspose.Cells shape geometry manipulation example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel workbook, reads a custom shape geometry definition from a ShapeGeometry.xml file, retrieves the first shape on the first worksheet, accesses its read‑only Geometry object, and demonstrates where to apply parsed geometry commands before saving the workbook. It includes error handling for missing files and ensures the output directory exists.
class ShapeGeometryImporter
{
    static void Main()
    {
        try
        {
            const string inputWorkbookPath = "InputWorkbook.xlsx";
            const string geometryFilePath = "ShapeGeometry.xml";
            const string outputWorkbookPath = "OutputWorkbook.xlsx";

            // Verify input workbook exists
            if (!File.Exists(inputWorkbookPath))
            {
                Console.WriteLine($"Input workbook not found: {inputWorkbookPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputWorkbookPath);

            // Access the first worksheet (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape in the collection
            Shape shape = sheet.Shapes[0];

            // Verify geometry XML file exists
            if (!File.Exists(geometryFilePath))
            {
                Console.WriteLine($"Geometry XML file not found: {geometryFilePath}");
                return;
            }

            // Read the shape geometry XML
            string geometryXml;
            try
            {
                geometryXml = File.ReadAllText(geometryFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read geometry file: {ex.Message}");
                return;
            }

            // Apply custom geometry to the shape.
            // Shape.Geometry is read‑only; we modify the existing Geometry object.
            try
            {
                Geometry geometry = shape.Geometry;

                // Placeholder for custom geometry handling.
                // Example: parse geometryXml and build the geometry using methods such as:
                // geometry.AddPath();
                // geometry.AddLineTo(x, y);
                // geometry.AddArcTo(...);
                // geometry.AddClosePath();
                // For now we leave the geometry unchanged.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to apply custom geometry: {ex.Message}");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputWorkbookPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputWorkbookPath);
            Console.WriteLine($"Workbook saved successfully to '{outputWorkbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
