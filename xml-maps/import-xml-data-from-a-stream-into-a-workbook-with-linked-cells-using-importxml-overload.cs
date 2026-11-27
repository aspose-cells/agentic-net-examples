// Title: Import XML data from a MemoryStream into an Aspose.Cells workbook using ImportXml with an XML map (C#)
// Common Searches: C# Aspose.Cells import XML from MemoryStream into worksheet using ImportXml overload | example adding XML map from XSD and loading XML stream into Excel with Aspose.Cells | import XML string as stream to linked cells in Aspose.Cells workbook C# | Aspose.Cells ImportXml method with stream parameter and start cell coordinates | save workbook after importing XML and export map back to XML using Aspose.Cells C#
// Tags: ImportXml from MemoryStream Aspose.Cells | XML map integration Aspose.Cells C# | load XSD schema and import XML Aspose.Cells | export XML from workbook map Aspose.Cells | populate worksheet cells from XML stream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ImportXmlFromStreamDemo
{
    // Demonstrates creating a workbook, adding an XML map from an XSD file, converting an XML string to a MemoryStream, importing the XML into the first worksheet at cell A1 with ImportXml, saving the workbook, exporting the data back to XML via the map, and printing sample cell values to verify the import.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for output files
            const string outputWorkbookPath = "ImportedXmlDemo.xlsx";
            const string exportedXmlPath = "ExportedFromMap.xml";
            const string xsdFilePath = "ProductsSchema.xsd";

            // XSD schema defining the XML structure
            string xsdSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                  <xs:element name='Products'>
                    <xs:complexType>
                      <xs:sequence>
                        <xs:element name='Product' maxOccurs='unbounded'>
                          <xs:complexType>
                            <xs:sequence>
                              <xs:element name='Name' type='xs:string'/>
                              <xs:element name='Price' type='xs:decimal'/>
                            </xs:sequence>
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>
                </xs:schema>";

            // Sample XML data that conforms to the schema
            string xmlData = @"
                <Products>
                  <Product>
                    <Name>Laptop</Name>
                    <Price>999.99</Price>
                  </Product>
                  <Product>
                    <Name>Phone</Name>
                    <Price>699.99</Price>
                  </Product>
                  <Product>
                    <Name>Tablet</Name>
                    <Price>399.99</Price>
                  </Product>
                </Products>";

            // Ensure the XSD file exists (write it if necessary)
            try
            {
                if (!File.Exists(xsdFilePath))
                {
                    File.WriteAllText(xsdFilePath, xsdSchema);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating XSD file: {ex.Message}");
                return;
            }

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Add the XML map to the workbook using the XSD file
            XmlMap xmlMap;
            try
            {
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xsdFilePath);
                xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                Console.WriteLine($"XmlMap added. Name: {xmlMap.Name}, Root element: {xmlMap.RootElementName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding XmlMap: {ex.Message}");
                return;
            }

            // Convert the XML string into a MemoryStream (required by ImportXml overload)
            using (MemoryStream xmlStream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(xmlStream))
                {
                    writer.Write(xmlData);
                    writer.Flush();
                    xmlStream.Position = 0; // Reset to beginning
                }

                // Import the XML data into the first worksheet starting at cell A1
                try
                {
                    workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
                    Console.WriteLine("XML data imported successfully into the workbook.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during ImportXml: {ex.Message}");
                    return;
                }
            }

            // Save the workbook
            try
            {
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputWorkbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
                return;
            }

            // Export the data back to XML using the map
            try
            {
                workbook.ExportXml(xmlMap.Name, exportedXmlPath);
                Console.WriteLine($"Exported XML from map saved to '{exportedXmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting XML: {ex.Message}");
            }

            // Display a few cell values to verify import
            try
            {
                Worksheet sheet = workbook.Worksheets["Sheet1"];
                Console.WriteLine("Sample imported cell values:");
                Console.WriteLine($"A1 (Header or first mapped cell): {sheet.Cells[0, 0].StringValue}");
                Console.WriteLine($"A2 (First product name): {sheet.Cells[1, 0].StringValue}");
                Console.WriteLine($"B2 (First product price): {sheet.Cells[1, 1].StringValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading cell values: {ex.Message}");
            }
        }
    }
}
