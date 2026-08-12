// Title: Link Excel cells to XML map elements with Aspose.Cells in C#
// Description: Shows how to create a workbook, import an XSD as an XML map, assign a friendly name, bind cells A1 and B1 to XML nodes using Cells.LinkToXmlMap, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# | XML map | LinkToXmlMap | XSD schema | bind cell to XML | import XML to Excel | workbook XML integration | .NET Excel API | XML data binding
// Common Searches: Aspose.Cells link cell to XML element C# | Add XML map from XSD using Aspose.Cells | Cells.LinkToXmlMap example | Bind Excel cell to XML with Aspose.Cells .NET | Programmatically import XML into Excel workbook
// Developer Intent: Add an XML map from an XSD file, connect specific worksheet cells to XML nodes, and generate an Excel file that stays synchronized with the XML source.
// Use Cases: Create a reporting template where cell values automatically reflect XML data fields, enabling live updates when the XML changes. | Build a round‑trip data exchange solution: users edit linked cells in Excel, and the underlying XML is updated on save. | Automate extraction of XML content into a spreadsheet for downstream analysis while preserving the original XML structure.
// AI Prompts: Generate C# code to load an existing XML file into the workbook after cells have been linked with Cells.LinkToXmlMap. | Show how to modify linked cell values programmatically and export the updated XML back to a file using Aspose.Cells. | Explain handling of XML namespaces in XPath expressions when using Cells.LinkToXmlMap.

using System;
using Aspose.Cells;

namespace AsposeCellsXmlLinkDemo
{
    // Shows how to create a workbook, import an XSD as an XML map, assign a friendly name, bind cells A1 and B1 to XML nodes using Cells.LinkToXmlMap, and save the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Path to the XML schema (XSD) that defines the XML structure.
            // The schema file should exist at the specified location.
            string schemaPath = "schema.xsd";

            // Add the XML map to the workbook using the schema file.
            // The Add method returns the index of the newly added map.
            int mapIndex = wb.Worksheets.XmlMaps.Add(schemaPath);

            // Retrieve the XmlMap object and give it a friendly name.
            XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "MyXmlMap";

            // Get the first worksheet and its cells collection.
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example: Link cell A1 (row 0, column 0) to an XML element.
            // Adjust the XPath to match the element you want to bind.
            string xmlElementPath = "/Root/Item/Name";
            cells.LinkToXmlMap(xmlMap.Name, 0, 0, xmlElementPath);

            // Example: Link cell B1 (row 0, column 1) to another XML element.
            string xmlElementPath2 = "/Root/Item/Price";
            cells.LinkToXmlMap(xmlMap.Name, 0, 1, xmlElementPath2);

            // Save the workbook with the linked XML map.
            wb.Save("LinkedXmlMapWorkbook.xlsx");
        }
    }
}
