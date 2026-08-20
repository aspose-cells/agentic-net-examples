// Title: Add an XML map from an XSD string and import XML data with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a Workbook, load an XSD schema directly from a string (no permanent file), add the XML map, import XML content from a memory stream, and save the result as an Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | XML map from string | XSD schema in memory | ImportXml | dynamic XML schema | memory stream Excel export | runtime XML mapping | no temporary file | Excel workbook generation
// Common Searches: Aspose.Cells add XML map from XSD string | Import XML into Excel without saving XSD file | C# load XML map from string Aspose.Cells | How to use ImportXml with a memory stream | Create XML map at runtime Aspose.Cells .NET
// Developer Intent: Create a workbook, add an XML map from an XSD string, import XML data, and save the Excel file.
// Use Cases: Generate XML maps on the fly from service‑provided XSD strings and map incoming XML data to Excel. | Process XML payloads from APIs without writing schema files to disk, improving security and performance. | Automate reporting by converting product catalogs or other XML datasets into formatted Excel worksheets.
// AI Prompts: Show C# code that adds an XML map to a Workbook directly from an XSD string without using a temporary file. | Provide an example of importing XML data from a string into a worksheet after creating the XML map from an XSD string. | Explain best practices for error handling and cleanup when using XmlMaps.Add with in‑memory schemas in Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsXmlMapFromString
{
    // Demonstrates how to create a Workbook, load an XSD schema directly from a string (no permanent file), add the XML map, import XML content from a memory stream, and save the result as an Excel file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // XML schema (XSD) as a string – defines the mapping structure
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
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

                // Sample XML data that conforms to the above schema
                string xmlData = @"<Products>
                                       <Product>
                                           <Name>Laptop</Name>
                                           <Price>999.99</Price>
                                       </Product>
                                       <Product>
                                           <Name>Phone</Name>
                                           <Price>699.99</Price>
                                       </Product>
                                   </Products>";

                // Create a new workbook (empty workbook)
                Workbook workbook = new Workbook();

                // Write the XSD string to a temporary file because Aspose.Cells expects a file path
                string tempXsdPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
                File.WriteAllText(tempXsdPath, xmlSchema, Encoding.UTF8);

                // Ensure the temporary XSD file exists before adding the XML map
                if (!File.Exists(tempXsdPath))
                    throw new FileNotFoundException("Temporary XSD file was not created.", tempXsdPath);

                // Add the XML map to the workbook using the temporary XSD file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "ProductsMap"; // optional: give the map a friendly name

                // Clean up the temporary XSD file (optional)
                try { File.Delete(tempXsdPath); } catch { /* ignore cleanup errors */ }

                // Import the XML data into the first worksheet starting at cell A1 (row 0, column 0)
                using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlData)))
                {
                    workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
                }

                // Save the workbook to an Excel file
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProductsMapped.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
