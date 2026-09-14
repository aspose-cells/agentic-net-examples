// Title: Validate SVG files for unsupported elements before inserting them as pictures into an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write a C# method that loads an SVG file, scans its XML for Aspose.Cells unsupported tags (script, foreignObject, animate, etc.), and returns a list of offending element names. | Show how to open an existing Excel workbook with Aspose.Cells, call the SVG validation method, and add the SVG as a picture only when no disallowed elements are detected. | Create logging code that records each unsupported SVG element found and aborts the picture insertion to avoid rendering failures.
// Common Searches: how to programmatically check an SVG for script and foreignObject tags before using Aspose.Cells | C# Aspose.Cells validate SVG elements that are not supported for picture insertion | prevent Aspose.Cells rendering errors by filtering unsupported SVG elements | example of inserting an SVG into Excel only after validation with Aspose.Cells .NET | detect and list disallowed SVG tags in a .NET application using Aspose.Cells
// Tags: svg validation with Aspose.Cells | detect unsupported svg elements in C# | insert svg picture into Excel using Aspose.Cells | prevent svg rendering errors in Aspose.Cells | filter script foreignObject tags from svg

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;

namespace SvgValidationExample
{
    // The program loads an SVG file, checks it for elements that Aspose.Cells cannot render (such as script, foreignObject, animate, etc.), reports any unsupported tags, and inserts the SVG as a picture into the first worksheet of an Excel workbook only when the validation passes.
    class Program
    {
        // List of SVG elements that Aspose.Cells does not support for rendering.
        private static readonly HashSet<string> UnsupportedElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "script",
            "foreignObject",
            "animate",
            "animateMotion",
            "animateTransform",
            "set",
            "metadata"
        };

        static void Main(string[] args)
        {
            try
            {
                // Path to the Excel file to work with.
                string workbookPath = "Sample.xlsx";

                // Path to the SVG file to be inserted.
                string svgPath = "Image.svg";

                // Verify that the workbook exists.
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: '{workbookPath}'.");
                    return;
                }

                // Verify that the SVG file exists before validation.
                if (!File.Exists(svgPath))
                {
                    Console.WriteLine($"SVG file not found: '{svgPath}'.");
                    return;
                }

                // Validate the SVG before insertion.
                if (!ValidateSvg(svgPath, out var unsupportedFound))
                {
                    Console.WriteLine("SVG validation failed. Unsupported elements detected:");
                    foreach (var elem in unsupportedFound)
                    {
                        Console.WriteLine($" - {elem}");
                    }
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(workbookPath);

                // Insert the SVG into the first worksheet at cell A1.
                Worksheet sheet = workbook.Worksheets[0];

                // Aspose.Cells can add pictures from a stream; the SVG will be rendered as an image.
                using (FileStream svgStream = new FileStream(svgPath, FileMode.Open, FileAccess.Read))
                {
                    // Add picture; Aspose.Cells will handle conversion if supported.
                    sheet.Pictures.Add(0, 0, svgStream);
                }

                // Save the workbook with the inserted SVG.
                string outputPath = "Sample_With_SVG.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="svgFilePath">Full path to the SVG file.</param>
        /// <param name="unsupportedElements">List of unsupported element names found.</param>
        /// <returns>True if the SVG does not contain unsupported elements; otherwise false.</returns>
        private static bool ValidateSvg(string svgFilePath, out List<string> unsupportedElements)
        {
            unsupportedElements = new List<string>();

            // File existence already checked by caller, but double‑check for safety.
            if (!File.Exists(svgFilePath))
            {
                throw new FileNotFoundException("SVG file not found.", svgFilePath);
            }

            XDocument svgDoc;
            try
            {
                svgDoc = XDocument.Load(svgFilePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to load SVG file as XML.", ex);
            }

            // Search for any unsupported elements in the SVG.
            var allElements = svgDoc.Descendants()
                                   .Select(e => e.Name.LocalName)
                                   .Distinct();

            foreach (var elem in allElements)
            {
                if (UnsupportedElements.Contains(elem))
                {
                    unsupportedElements.Add(elem);
                }
            }

            // Return true if no unsupported elements were found.
            return unsupportedElements.Count == 0;
        }
    }
}
