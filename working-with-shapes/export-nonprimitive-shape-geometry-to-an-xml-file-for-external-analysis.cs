// Title: Export Excel worksheet shape geometry to XML using Aspose.Cells for .NET (C#)
// Common Searches: how to extract shape coordinates from an Excel file with Aspose.Cells C# | Aspose.Cells C# export shape size and position to XML | save Excel drawing objects details to XML using Aspose.Cells | C# code to list all shapes in a workbook and write to XML | export non‑primitive shape properties from .xlsx to XML with Aspose.Cells
// Tags: Aspose.Cells export shape metadata to XML | C# extract Excel shape attributes | worksheet shape geometry serialization | Aspose.Cells shape collection iteration | XML generation for Excel drawing objects

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Xml.Linq;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and its Shapes collection, captures each shape's Name, Type, Top, Left, Width, and Height into XML elements, aggregates them under a <Shapes> root node, and saves the resulting XML document. It includes error handling for missing files and individual shape processing failures.
class ExportShapeGeometry
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "ShapeGeometry.xml";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Root element for the XML document
            XElement root = new XElement("Shapes");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Create an XML element for the shape with basic attributes
                        XElement shapeElement = new XElement("Shape",
                            new XAttribute("Name", shape.Name),
                            new XAttribute("Type", shape.Type.ToString()),
                            new XAttribute("Top", shape.Top),
                            new XAttribute("Left", shape.Left),
                            new XAttribute("Width", shape.Width),
                            new XAttribute("Height", shape.Height));

                        // Add the shape element to the root
                        root.Add(shapeElement);
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Warning: Could not process shape \"{shape.Name}\". {shapeEx.Message}");
                    }
                }
            }

            // Build the final XML document and save it
            XDocument xmlDoc = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                root);
            xmlDoc.Save(outputPath);
            Console.WriteLine($"Shape information exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
