// Title: C# – Load an .xlsx Workbook and Attach an XML Schema (XSD) Map with Aspose.Cells
// Description: Demonstrates how to open an existing .xlsx file using Aspose.Cells for .NET, add an XML schema (XSD) as an XmlMap via Worksheets.XmlMaps.Add, optionally assign a friendly name, and save the workbook with the map embedded. Includes tips for retrieving the map index and handling missing schema files.
// Keywords: Aspose.Cells C# load workbook | add XML map Aspose.Cells | attach XSD schema to Excel | Worksheets.XmlMaps.Add | set XmlMap name | save workbook with XmlMap | .NET Excel XML mapping | C# XML schema map example
// Common Searches: add xml map to existing excel file aspose.cells | c# load xlsx and attach xsd schema | aspnet aspose cells xml map tutorial | how to embed xml schema in excel using c# | aspose.cells Worksheets.XmlMaps example
// Developer Intent: Open a .xlsx file, add an XSD as an XmlMap, optionally rename it, and save the updated workbook.
// Use Cases: Create a template workbook pre‑linked to an XSD for automated XML import/export. | Distribute a spreadsheet with a named XmlMap so downstream users can map XML data without manual setup. | Integrate Excel‑XML round‑trip in a data‑exchange pipeline by embedding the schema directly in the workbook.
// AI Prompts: Write C# code that opens an .xlsx workbook, adds an XSD file as an XmlMap, sets a custom map name, and saves the file using Aspose.Cells. | Explain how to retrieve the index of a newly added XmlMap and handle missing or invalid XSD files in Aspose.Cells. | Provide a guide to export workbook data back to XML using the attached XmlMap with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to open an existing .xlsx file using Aspose.Cells for .NET, add an XML schema (XSD) as an XmlMap via Worksheets.XmlMaps.Add, optionally assign a friendly name, and save the workbook with the map embedded. Includes tips for retrieving the map index and handling missing schema files.
class Program
{
    static void Main()
    {
        // Path to the existing Excel workbook (.xlsx)
        string excelFilePath = "input.xlsx";

        // Load the workbook using the constructor that accepts a file path
        Workbook workbook = new Workbook(excelFilePath);

        // Path to the XML schema (XSD) that defines the XML map
        string xmlSchemaPath = "schema.xsd";

        // Add the XML schema as a map to the workbook
        // The Add method returns the index of the newly added XmlMap
        int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchemaPath);

        // Retrieve the XmlMap object (optional, e.g., to set a friendly name)
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "MyXmlMap";

        // Save the workbook with the attached XML map
        workbook.Save("output_with_xmlmap.xlsx");
    }
}
