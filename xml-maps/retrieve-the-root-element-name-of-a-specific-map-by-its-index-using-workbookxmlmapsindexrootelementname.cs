// Title: Aspose.Cells .NET – Get XML Map Root Element Name by Index
// Description: Shows how to add an XML schema to a workbook, capture the map index, and read the map’s RootElementName using workbook.Worksheets.XmlMaps[index].RootElementName, with optional saving of the workbook.
// Keywords: Aspose.Cells | .NET | C# | XmlMaps | RootElementName | XML map index | retrieve root element | Aspose.Cells XML mapping | get schema root name | Workbook XmlMaps
// Common Searches: Aspose.Cells get XML map root element name by index | C# retrieve RootElementName from XmlMaps collection | How to read XML map schema root in Aspose.Cells | XmlMaps index access example Aspose.Cells | Get root element of XML map in .NET workbook
// Developer Intent: Obtain the root element name of a specific XML map using its numeric index.
// Use Cases: Confirm that an imported XML map matches the expected schema before data import. | Log or display root element names of all XML maps in a workbook for debugging. | Drive conditional processing when multiple XML maps are present by checking each map’s RootElementName.
// AI Prompts: Write C# code that iterates through all XmlMaps in a workbook and prints each RootElementName. | Create a method that returns the RootElementName for a given XmlMap index and includes error handling for invalid indexes. | Provide an example that adds several XML schemas to a workbook and retrieves the root element name of each map.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to add an XML schema to a workbook, capture the map index, and read the map’s RootElementName using workbook.Worksheets.XmlMaps[index].RootElementName, with optional saving of the workbook.
    public class RetrieveXmlMapRootElementDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample XML schema (XSD) defining a root element named "Data"
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Data'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string' />
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                // Add the XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);

                // Retrieve the root element name of the added map by its index
                string rootElementName = workbook.Worksheets.XmlMaps[mapIndex].RootElementName;

                // Display the result
                Console.WriteLine($"Root Element Name of map at index {mapIndex}: {rootElementName}");

                // Save the workbook (optional, just to demonstrate lifecycle compliance)
                workbook.Save("RetrieveXmlMapRootElementDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveXmlMapRootElementDemo.Run();
        }
    }
}
