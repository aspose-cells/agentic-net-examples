// Title: Validate SVG for Unsupported Elements Before Adding to Aspose.Cells Worksheet (C#)
// Description: C# example that loads an SVG, parses its XML, checks for Aspose.Cells unsupported tags (script, foreignObject, animate, etc.), and inserts the graphic only when validation succeeds, preventing rendering errors.
// Keywords: Aspose.Cells SVG validation | C# SVG parsing | unsupported SVG tags | AddSvg ShapeCollection | script tag detection | foreignObject check | Excel workbook SVG | XML validation C# | Aspose.Cells supported elements | batch SVG processing
// Common Searches: Aspose.Cells validate SVG before AddSvg | C# check unsupported SVG tags Aspose | prevent script tag errors in Aspose.Cells SVG | list of SVG elements supported by Aspose.Cells | validate SVG file for Excel insertion C#
// Developer Intent: Ensure an SVG does not contain tags that Aspose.Cells cannot render before inserting it into a worksheet.
// Use Cases: Pre‑process user‑uploaded SVGs to avoid rendering failures in generated Excel files. | Automate batch validation of SVG assets before creating reports with multiple worksheets. | Log offending element names and skip files that contain disallowed tags. | Extend the check to include size limits or empty‑content detection.
// AI Prompts: Write a C# method that accepts an SVG byte array, returns a bool and a list of Aspose.Cells unsupported element names with a friendly validation message. | Create a C# console application that scans a folder of SVG files, validates each using the provided logic, and adds only the valid graphics to separate worksheets in a new workbook, including error handling and logging.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSvgValidation
{
    // C# example that loads an SVG, parses its XML, checks for Aspose.Cells unsupported tags (script, foreignObject, animate, etc.), and inserts the graphic only when validation succeeds, preventing rendering errors.
    class Program
    {
        // List of SVG elements that Aspose.Cells does not support and may cause rendering errors
        private static readonly HashSet<string> UnsupportedSvgElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "script",
            "foreignObject",
            "animate",
            "set",
            "animateMotion",
            "animateTransform",
            "animateColor"
        };

        static void Main()
        {
            try
            {
                // Path to the SVG file to be inserted
                const string svgPath = "image.svg";

                // Verify that the SVG file exists before attempting to read it
                if (!File.Exists(svgPath))
                {
                    Console.WriteLine($"SVG file not found: {svgPath}");
                    return;
                }

                // Load the SVG file into a byte array
                byte[] svgData = File.ReadAllBytes(svgPath);

                // Validate the SVG content before adding it to the worksheet
                if (!IsSvgSupported(svgData, out string validationMessage))
                {
                    Console.WriteLine("SVG validation failed: " + validationMessage);
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                ShapeCollection shapes = worksheet.Shapes;

                // Add the validated SVG to the worksheet.
                // Using rows 4‑5 and columns 5‑10 as an example area; Aspose.Cells will size the shape within this range.
                // Offsets (0,0) are used to position the shape at the top‑left corner of the specified range.
                shapes.AddSvg(4, 5, 10, 10, 0, 0, svgData, null);

                // Save the workbook
                const string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully with validated SVG at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }

        /// <param name="svgBytes">Raw SVG file bytes.</param>
        /// <param name="message">Detailed validation message.</param>
        /// <returns>True if SVG is supported; otherwise false.</returns>
        private static bool IsSvgSupported(byte[] svgBytes, out string message)
        {
            try
            {
                // Load SVG XML from the byte array
                XDocument doc;
                using (MemoryStream ms = new MemoryStream(svgBytes))
                {
                    doc = XDocument.Load(ms);
                }

                // Search for any unsupported elements in the document
                var found = doc.Descendants()
                               .Where(e => UnsupportedSvgElements.Contains(e.Name.LocalName))
                               .Select(e => e.Name.LocalName)
                               .Distinct()
                               .ToList();

                if (found.Any())
                {
                    message = "Unsupported SVG elements detected: " + string.Join(", ", found);
                    return false;
                }

                // Additional optional checks (e.g., empty SVG) can be added here

                message = "SVG is valid.";
                return true;
            }
            catch (Exception ex)
            {
                // XML parsing errors indicate an invalid SVG file
                message = "Error parsing SVG: " + ex.Message;
                return false;
            }
        }
    }
}
