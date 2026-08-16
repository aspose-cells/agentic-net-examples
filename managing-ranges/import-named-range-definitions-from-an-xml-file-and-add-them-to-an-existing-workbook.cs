// Title: Import Named Ranges from XML into an Existing Excel Workbook with Aspose.Cells (C#)
// Description: C# sample that validates the presence of a workbook and an XML file, loads the workbook using Aspose.Cells, parses <NamedRange> elements, adds each name to the Worksheets.Names collection, assigns the RefersTo formula, and saves the updated file.
// Keywords: Aspose.Cells import named ranges XML | C# add named range to workbook | load named range definitions from XML | Workbook.Worksheets.Names.Add example | read XML named ranges C# | Excel named range automation | Aspose.Cells XML parsing | programmatic named range creation | .NET Excel named range import
// Common Searches: how to import named ranges from xml using aspose.cells | c# add named range to existing workbook from xml | aspose.cells read named range definitions xml | load custom named ranges into Excel with .net | sample code for xml based named range import aspnet
// Developer Intent: Read named‑range definitions stored in an XML file and programmatically add them to an existing Excel workbook using Aspose.Cells for .NET.
// Use Cases: Migrate legacy XML‑defined named ranges into newly generated reports. | Synchronize a shared set of named ranges across multiple workbooks by importing a central XML file. | Prepare workbooks automatically before applying formulas, charts, or data‑validation that rely on predefined named ranges.
// AI Prompts: Write C# code that parses an XML file containing <NamedRange> elements and adds each as a named range to an Aspose.Cells Workbook, with robust error handling for missing attributes. | Refactor the example to use LINQ to XML for reading named range definitions and add logging that records each successful import and any failures. | Explain how to extend the import routine to support sheet‑scoped named ranges (e.g., Name="MyRange" Sheet="Sheet2") when reading from the XML.

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

// C# sample that validates the presence of a workbook and an XML file, loads the workbook using Aspose.Cells, parses <NamedRange> elements, adds each name to the Worksheets.Names collection, assigns the RefersTo formula, and saves the updated file.
class ImportNamedRangesFromXml
{
    static void Main()
    {
        // Paths to the existing workbook and the XML file that contains named range definitions
        string workbookPath = "ExistingWorkbook.xlsx";
        string xmlPath = "NamedRanges.xml";

        try
        {
            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

            // Load the existing workbook
            Workbook wb = new Workbook(workbookPath);

            // Verify that the XML file exists
            if (!File.Exists(xmlPath))
                throw new FileNotFoundException($"XML file not found: {xmlPath}");

            // Load and parse the XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);

            // Expected XML format:
            // <NamedRanges>
            //   <NamedRange Name="MyRange" RefersTo="=Sheet1!$A$1:$B$2" />
            //   ...
            // </NamedRanges>

            XmlNodeList rangeNodes = doc.SelectNodes("//NamedRange");
            if (rangeNodes != null)
            {
                foreach (XmlNode node in rangeNodes)
                {
                    // Retrieve the name and reference from attributes
                    string name = node.Attributes["Name"]?.Value;
                    string refersTo = node.Attributes["RefersTo"]?.Value;

                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(refersTo))
                    {
                        // Add the named range to the workbook's Names collection
                        int index = wb.Worksheets.Names.Add(name);
                        wb.Worksheets.Names[index].RefersTo = refersTo;
                    }
                }
            }

            // Save the updated workbook
            string outputPath = "WorkbookWithImportedNames.xlsx";
            wb.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
