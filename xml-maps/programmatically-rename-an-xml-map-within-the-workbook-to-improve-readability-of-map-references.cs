// Title: C# Example: Rename an XML Map in an Aspose.Cells Workbook and Export XML
// Description: Demonstrates how to create a workbook with Aspose.Cells for .NET, add an XML schema, insert an XML map, rename the map using the XmlMap.Name property, export the data to an XML file with the new map name, and save the workbook.
// Keywords: Aspose.Cells XML map rename C# | XmlMap.Name property | ExportXml with custom map name | .NET workbook XML mapping | programmatically rename XML map | Aspose.Cells example GitHub | C# Excel XML map tutorial
// Common Searches: rename XML map Aspose.Cells C# | how to change XmlMap name in .NET | export XML using renamed map Aspose.Cells | Aspose.Cells XML map example code | C# Aspose.Cells rename map before export
// Developer Intent: Change the name of an existing XML map in a workbook to a meaningful identifier and then generate the XML output using that new name.
// Use Cases: Standardize map names (e.g., "EmployeeMap") after adding a schema to improve code readability. | Align XML map identifiers with business terminology before sending data to downstream systems. | Allow end‑users to specify a custom map name at runtime and export the corresponding XML file.
// AI Prompts: Generate C# code with Aspose.Cells that adds an XML schema, creates an XML map, renames it to a custom name, and exports the XML. | Explain the purpose of the XmlMap.Name property in Aspose.Cells and show how it affects ExportXml. | Provide a step‑by‑step tutorial for renaming an XML map in a workbook and exporting the data, suitable for inclusion in a GitHub repository.

using System;
using System.IO;
using Aspose.Cells;

namespace Demo
{
    // Demonstrates how to create a workbook with Aspose.Cells for .NET, add an XML schema, insert an XML map, rename the map using the XmlMap.Name property, export the data to an XML file with the new map name, and save the workbook.
    class RenameXmlMapDemo
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["B2"].PutValue("John Doe");

            // Define an XML schema for the map
            string xmlSchema = "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">" +
                               "<xs:element name=\"Employee\">" +
                               "<xs:complexType><xs:sequence>" +
                               "<xs:element name=\"ID\" type=\"xs:int\"/>" +
                               "<xs:element name=\"Name\" type=\"xs:string\"/>" +
                               "</xs:sequence></xs:complexType></xs:element></xs:schema>";

            // Add the XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Rename the XML map
            xmlMap.Name = "EmployeeMap";

            // Export XML using the new map name
            string xmlPath = "EmployeeData.xml";
            workbook.ExportXml(xmlMap.Name, xmlPath);
            Console.WriteLine($"XML exported to: {Path.GetFullPath(xmlPath)}");

            // Save the workbook
            string workbookPath = "RenamedXmlMapWorkbook.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
        }
    }
}
