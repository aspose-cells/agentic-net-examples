// Title: Validate SVG for Unsupported Elements Before Adding to Aspose.Cells Worksheet (C#)
// Description: Loads an SVG file into a byte array, parses it with XDocument, and checks for disallowed tags (script, foreignObject, iframe, object, embed). If the SVG passes the check, it is inserted into a worksheet via ShapeCollection.AddSvg and the workbook is saved.
// Keywords: Aspose.Cells SVG validation | C# SVG unsupported tags | insert SVG shape Excel | filter script tag Aspose.Cells | validate SVG before insertion | Excel shape SVG C# | Aspose.Cells AddSvg example
// Common Searches: how to check SVG for unsupported elements in Aspose.Cells | C# code to prevent script tags in SVG when adding to Excel | validate SVG before using ShapeCollection.AddSvg | Aspose.Cells SVG rendering errors cause | skip SVG with foreignObject in Aspose.Cells
// Developer Intent: Ensure an SVG file contains no elements that Aspose.Cells cannot render before inserting it as a shape in a worksheet.
// Use Cases: Batch‑process a folder of SVGs, inserting only those that pass validation to avoid runtime exceptions. | Log filenames of rejected SVGs and continue processing the remaining files in an automated report generator. | Display a clear warning to end‑users when they select an SVG that includes script, foreignObject, iframe, object, or embed tags.
// AI Prompts: Write a C# method that receives an SVG byte array and returns true only if it lacks script, foreignObject, iframe, object, and embed elements. | Create error‑handling logic for SVG validation in Aspose.Cells that logs failures and shows user‑friendly messages. | Generate unit tests for IsSvgSupported covering SVGs with only supported elements and SVGs containing each prohibited tag.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an SVG file into a byte array, parses it with XDocument, and checks for disallowed tags (script, foreignObject, iframe, object, embed). If the SVG passes the check, it is inserted into a worksheet via ShapeCollection.AddSvg and the workbook is saved.
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

    // Checks whether the SVG byte array contains any unsupported elements
    static bool IsSvgSupported(byte[] svgData)
    {
        try
        {
            // Load SVG XML from the byte array
            XDocument doc = XDocument.Load(new MemoryStream(svgData));

            // Search for any disallowed element names (case‑insensitive)
            var badElements = doc.Descendants()
                                 .Where(e => UnsupportedElements.Contains(e.Name.LocalName,
                                                                          StringComparer.OrdinalIgnoreCase));

            // If any are found, the SVG is not supported
            return !badElements.Any();
        }
        catch
        {
            // Parsing errors also mean the SVG is not suitable
            return false;
        }
    }

    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Path to the SVG file to be inserted
            string svgPath = "image.svg";

            // Ensure the SVG file exists before attempting to read it
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
                Console.WriteLine("The SVG file contains unsupported elements and will not be added.");
                return;
            }

            // Insert the validated SVG into the worksheet.
            // Using -1 for height and width lets Excel auto‑size the shape.
            shapes.AddSvg(4, 0, 5, 0, -1, -1, svgBytes, null);

            // Save the workbook with the inserted SVG
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully with validated SVG to {outputPath}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
