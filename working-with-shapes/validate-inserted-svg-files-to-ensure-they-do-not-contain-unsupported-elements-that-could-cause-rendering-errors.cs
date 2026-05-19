using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class SvgValidator
{
    // List of SVG elements that Aspose.Cells does not support
    static readonly string[] UnsupportedElements = new[]
    {
        "script",
        "foreignObject",
        "iframe",
        "object",
        "embed"
    };

    // Checks the SVG byte array for unsupported elements
    static bool IsSvgSupported(byte[] svgData)
    {
        try
        {
            // Load SVG XML from memory
            XDocument doc = XDocument.Load(new MemoryStream(svgData));

            // Get all element names in the document
            var elementNames = doc.Descendants()
                                  .Select(e => e.Name.LocalName);

            // Return false if any unsupported element is found
            return !elementNames.Any(name =>
                UnsupportedElements.Contains(name, StringComparer.OrdinalIgnoreCase));
        }
        catch
        {
            // If the SVG cannot be parsed, treat it as unsupported
            return false;
        }
    }

    static void Main()
    {
        try
        {
            // Path to the SVG file to be inserted
            string svgPath = "image.svg";

            // Verify that the SVG file exists
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Read the SVG file into a byte array
            byte[] svgBytes = File.ReadAllBytes(svgPath);

            // Validate the SVG content before insertion
            if (!IsSvgSupported(svgBytes))
            {
                Console.WriteLine("The SVG file contains unsupported elements and cannot be added.");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Insert the validated SVG into the worksheet
            // Overload requires picture width and height before the data arrays.
            // Using 0 for width/height lets Aspose.Cells calculate the size automatically.
            Picture picture = shapes.AddSvg(4, 5, 6, 6, 0, 0, svgBytes, null);

            // Save the workbook with the inserted SVG
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully with validated SVG: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}