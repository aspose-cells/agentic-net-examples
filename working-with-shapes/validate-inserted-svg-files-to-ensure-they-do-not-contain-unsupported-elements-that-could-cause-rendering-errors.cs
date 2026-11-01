// Title: How to validate SVG files for unsupported elements before inserting them as pictures with Aspose.Cells in C#
// AI Prompts: Write a C# method that loads an SVG document, iterates over every node, and returns a list of element names that are not included in Aspose.Cells' supported SVG element set. | Show the code that calls the SVG validator, checks the returned list, and only calls Worksheets.Pictures.Add when the SVG passes validation. | Extend the validator to treat <script> and <foreignObject> tags as unsupported and abort the insertion, then log the offending tags.
// Common Searches: C# validate SVG elements against Aspose.Cells supported list before adding picture | how to reject SVG files with disallowed tags when using Aspose.Cells Pictures.Add | list unsupported SVG nodes for Excel export using Aspose.Cells API | filter out script and foreignObject elements from SVG before inserting into workbook
// Tags: Aspose.Cells SVG element whitelist validation | C# filter unsupported SVG nodes before picture insertion | validate SVG compatibility with Aspose.Cells Pictures.Add | SVG rendering support list Aspose.Cells | detect disallowed SVG tags in C# for Excel export

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

// Provides a C# example that checks an SVG file against a predefined set of elements supported by Aspose.Cells, reports any unsupported tags, and inserts the SVG into a workbook only when validation succeeds.
class SvgValidator
{
    // List of SVG elements that Aspose.Cells can render
    private static readonly HashSet<string> SupportedElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "svg","g","path","rect","circle","ellipse","line","polyline","polygon",
        "text","tspan","defs","use","image","clipPath","mask","filter",
        "linearGradient","radialGradient","stop","style","metadata"
    };

    // Returns true if the SVG contains only supported elements
    public static bool ValidateSvg(string svgFilePath, out List<string> unsupportedElements)
    {
        unsupportedElements = new List<string>();

        if (!File.Exists(svgFilePath))
        {
            Console.WriteLine($"SVG file not found: {svgFilePath}");
            return false;
        }

        try
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(svgFilePath);

            // Select all elements in the document
            XmlNodeList allElements = doc.SelectNodes("//*");
            foreach (XmlNode node in allElements)
            {
                if (!SupportedElements.Contains(node.Name) && !unsupportedElements.Contains(node.Name))
                {
                    unsupportedElements.Add(node.Name);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while validating SVG: {ex.Message}");
            return false;
        }

        return unsupportedElements.Count == 0;
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Path to the SVG file to be inserted
            string svgPath = "example.svg";

            // Validate the SVG before insertion
            if (SvgValidator.ValidateSvg(svgPath, out var unsupported))
            {
                // Insert the SVG as a picture if it passed validation
                // (row index, column index, file path)
                workbook.Worksheets[0].Pictures.Add(0, 0, svgPath);
            }
            else
            {
                Console.WriteLine("SVG contains unsupported elements or could not be processed:");
                foreach (var elem in unsupported)
                {
                    Console.WriteLine("- " + elem);
                }
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
