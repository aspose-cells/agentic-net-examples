// Title: How to import shape geometry from an XML file and assign it to an existing worksheet shape using Aspose.Cells for .NET
// Common Searches: aspnet import custom shape geometry from xml into Excel worksheet using Aspose.Cells | set custom geometry of an Excel shape programmatically with Aspose.Cells C# | apply external shape geometry XML to a shape in a workbook via reflection Aspose.Cells | load shape geometry XML file and update shape in existing Excel file using Aspose.Cells | Aspose.Cells C# read shape geometry XML and modify shape properties
// Tags: import shape geometry xml Aspose.Cells | assign custom geometry to Excel shape | use reflection to set CustomGeometry property | update shape geometry in existing workbook | load external geometry XML for worksheet shape

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, reads a custom geometry XML file, uses reflection to set the shape's CustomGeometry property when supported, and saves the workbook with the modified shape.
class ShapeGeometryImporter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string geometryPath = "shapeGeometry.xml";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Workbook file '{inputPath}' not found.");
                return;
            }

            // Verify that the geometry XML file exists
            if (!File.Exists(geometryPath))
            {
                Console.WriteLine($"Error: Geometry file '{geometryPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("Error: No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape (or modify to select by name)
            Shape shape = sheet.Shapes[0];

            // Read the custom geometry XML from the external file
            string geometryXml = File.ReadAllText(geometryPath);

            // Attempt to set custom geometry via reflection (property may not exist in older versions)
            try
            {
                PropertyInfo customGeomProp = shape.GetType().GetProperty("CustomGeometry", BindingFlags.Public | BindingFlags.Instance);
                if (customGeomProp != null && customGeomProp.CanWrite)
                {
                    customGeomProp.SetValue(shape, geometryXml);
                    Console.WriteLine("Custom geometry applied to the shape.");
                }
                else
                {
                    Console.WriteLine("Warning: 'CustomGeometry' property not available in this Aspose.Cells version.");
                }
            }
            catch (Exception reflEx)
            {
                Console.WriteLine($"Error applying custom geometry via reflection: {reflEx.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
