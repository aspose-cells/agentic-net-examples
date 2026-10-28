// Title: Export all worksheet shape definitions to an XML file using Aspose.Cells for .NET
// Common Searches: how to extract shape information from an Excel workbook to XML with Aspose.Cells C# | Aspose.Cells C# export worksheet shapes metadata to external XML file | save shape name and type from Excel sheets as XML using Aspose.Cells | C# code to iterate workbook shapes and write definitions to XML for version control | Aspose.Cells example for exporting shape attributes to a separate XML document
// Tags: Aspose.Cells shape metadata XML export | C# iterate Excel shapes Aspose.Cells | export worksheet shape definitions to XML | track shape attributes version control Aspose.Cells | generate external XML for Excel shape objects

using System;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, loops through each worksheet and its Shapes collection, creates an XML document with a <Shapes> root, adds a <Shape> element for every shape containing SheetName, Name, and Type attributes, and saves the result to a separate XML file while handling missing files and shape‑specific errors.
class ExportShapesXml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "ShapesDefinitions.xml";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create an XML document that will hold all shape definitions
            XmlDocument shapesDoc = new XmlDocument();
            XmlElement rootElement = shapesDoc.CreateElement("Shapes");
            shapesDoc.AppendChild(rootElement);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Create an element representing the shape
                        XmlElement shapeElement = shapesDoc.CreateElement("Shape");

                        // Add identification attributes
                        shapeElement.SetAttribute("SheetName", sheet.Name);
                        shapeElement.SetAttribute("Name", shape.Name ?? string.Empty);
                        shapeElement.SetAttribute("Type", shape.Type.ToString());

                        // Append the shape element to the root
                        rootElement.AppendChild(shapeElement);
                    }
                    catch (Exception shapeEx)
                    {
                        // Log shape‑specific errors but continue processing other shapes
                        Console.WriteLine($"Warning: Could not process a shape on sheet \"{sheet.Name}\": {shapeEx.Message}");
                    }
                }
            }

            // Save the combined XML definitions to an external file
            shapesDoc.Save(outputPath);
            Console.WriteLine($"Shape definitions have been saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
