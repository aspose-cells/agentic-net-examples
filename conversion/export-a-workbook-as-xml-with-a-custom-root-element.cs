// Title: C# – Export Aspose.Cells Workbook to XML with a Custom Root Element via XSD Map
// Description: Creates a workbook, defines an XSD schema whose root is <Company>, adds the schema as an XmlMap, and exports the data to an XML file so the output starts with the custom root tag.
// Keywords: Aspose.Cells export XML | C# XML map XSD | custom root element | Workbook.ExportXml | Excel to XML schema | Aspose.Cells XmlMap example
// Common Searches: Aspose.Cells export workbook to XML with custom root | C# create XML map from XSD in Aspose.Cells | How to set root element when exporting Excel to XML | Export Excel data as XML using Aspose.Cells and XSD
// Developer Intent: Generate an XML file from a workbook where the root node is defined by an XSD schema rather than the default worksheet name.
// Use Cases: Produce an employee feed that conforms to a company‑wide XML schema for ERP integration. | Create XML configuration files with a specific root tag required by third‑party web services. | Automate data exchange between Excel and systems that expect a predefined XML structure.
// AI Prompts: Add a Department column to the worksheet and update the XSD so the exported XML includes <Department> under each <Employee>. | Load the XSD from an embedded resource instead of writing a temporary file before creating the XmlMap. | Export two worksheets to separate XML files, each using a different XmlMap with distinct root elements.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, defines an XSD schema whose root is <Company>, adds the schema as an XmlMap, and exports the data to an XML file so the output starts with the custom root tag.
class ExportWorkbookWithCustomRoot
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Employees";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define an XML schema whose root element is <Company>
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                    <xs:element name='Company'>
                        <xs:complexType>
                            <xs:sequence>
                                <xs:element name='Employee' maxOccurs='unbounded'>
                                    <xs:complexType>
                                        <xs:sequence>
                                            <xs:element name='Id' type='xs:int'/>
                                            <xs:element name='Name' type='xs:string'/>
                                        </xs:sequence>
                                    </xs:complexType>
                                </xs:element>
                            </xs:sequence>
                        </xs:complexType>
                    </xs:element>
                </xs:schema>";

            // Write the schema to a temporary file (required by Aspose.Cells API)
            string tempSchemaPath = Path.Combine(Path.GetTempPath(), "CompanySchema.xsd");
            File.WriteAllText(tempSchemaPath, xmlSchema);

            // Add the XML map to the workbook using the schema file
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "CompanyMap"; // optional custom map name

            // Export the workbook to XML using the map; the output will have <Company> as the root element
            string outputPath = "CompanyExport.xml";
            workbook.ExportXml(xmlMap.Name, outputPath);

            Console.WriteLine($"Workbook exported to XML with custom root element: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up temporary schema file if it exists
            try
            {
                string tempSchemaPath = Path.Combine(Path.GetTempPath(), "CompanySchema.xsd");
                if (File.Exists(tempSchemaPath))
                {
                    File.Delete(tempSchemaPath);
                }
            }
            catch
            {
                // Suppress any cleanup exceptions
            }
        }
    }
}
