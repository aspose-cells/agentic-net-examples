// Title: Generate an XML file listing all named ranges and their RefersTo formulas from an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to list all defined names in an Excel file with Aspose.Cells C# | export Excel named ranges to XML using Aspose.Cells .NET | C# Aspose.Cells create XML of workbook named ranges | retrieve RefersTo property of named ranges with Aspose.Cells | save Excel named range definitions to an XML document in C#
// Tags: Aspose.Cells named range XML export | C# extract workbook defined names | Aspose.Cells RefersTo extraction | Excel named ranges to XML conversion | Aspose.Cells generate XDocument for named ranges

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, iterates through workbook.Worksheets.Names, creates an XML <NamedRange> element for each defined name containing Name and RefersTo attributes, aggregates them under a <NamedRanges> root, and saves the XDocument to a specified file while handling missing files and Aspose.Cells exceptions.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "namedRanges.xml";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (no password supplied)
            Workbook workbook = new Workbook(inputPath);

            // Create the root element for the XML document
            XElement root = new XElement("NamedRanges");

            // Iterate through all defined names (named ranges) in the workbook
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                string nameText = definedName.Text;
                string refersTo = definedName.RefersTo;

                // Build an XML element for the current named range
                XElement rangeElement = new XElement("NamedRange",
                    new XAttribute("Name", nameText),
                    new XAttribute("RefersTo", refersTo));

                root.Add(rangeElement);
            }

            // Create the XDocument with declaration and root element
            XDocument doc = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                root);

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the XML representation to a file
            doc.Save(outputPath);
            Console.WriteLine($"Named ranges saved to {outputPath}");
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors (e.g., password protected)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
