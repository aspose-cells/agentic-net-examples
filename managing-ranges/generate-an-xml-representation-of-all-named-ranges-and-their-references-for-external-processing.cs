// Title: Create an XML file that lists all named ranges and their reference formulas from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, iterates through workbook.Worksheets.Names, and builds an XML document containing each named range's Name and RefersTo attributes. | Extend the XML output to include the worksheet name where each named range is defined, extracting the sheet information via Aspose.Cells APIs. | Add comprehensive error handling to the program to manage missing input files, empty named‑range collections, and failures when writing the XML document.
// Common Searches: asp.net c# export excel named ranges to xml using aspose.cells | how to list defined names from an xlsx file with Aspose.Cells .NET | generate xml of named range references from workbook using Aspose.Cells C# | c# retrieve RefersTo formula for each named range in Excel with Aspose.Cells | save named ranges and their sheet references to an xml file in .NET
// Tags: export named ranges to XML Aspose.Cells | enumerate defined names workbook Aspose.Cells | retrieve RefersTo formula C# | Aspose.Cells generate XML document | C# list Excel named ranges | handle missing workbook file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Xml.Linq;

// The program loads an Excel workbook via Aspose.Cells, iterates over all defined names, captures each named range's name and RefersTo formula, creates an XML document with these details, and saves it as NamedRanges.xml.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Create the root element for the XML document
        XElement root = new XElement("NamedRanges");

        // Iterate through all defined names (named ranges) in the workbook
        foreach (Name definedName in workbook.Worksheets.Names)
        {
            // Name of the range
            string name = definedName.Text;

            // Reference formula (e.g., =Sheet1!$A$1:$B$2)
            string refersTo = definedName.RefersTo;

            // Build an XML element for this named range
            XElement rangeElement = new XElement("NamedRange",
                new XAttribute("Name", name),
                new XAttribute("RefersTo", refersTo));

            // Add the element to the root
            root.Add(rangeElement);
        }

        // Assemble the complete XML document
        XDocument xmlDoc = new XDocument(
            new XDeclaration("1.0", "utf-8", "yes"),
            root);

        // Save the XML to a file (replace with desired output path)
        string xmlOutputPath = "NamedRanges.xml";
        xmlDoc.Save(xmlOutputPath);

        Console.WriteLine($"XML file with named ranges saved to: {Path.GetFullPath(xmlOutputPath)}");
    }
}
