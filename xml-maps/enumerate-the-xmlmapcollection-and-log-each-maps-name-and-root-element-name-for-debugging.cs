// Title: C# – Enumerate XmlMapCollection and Log Map Name & RootElementName with Aspose.Cells
// Description: Demonstrates how to create a workbook, add an XML schema as an XmlMap, retrieve the workbook's XmlMapCollection, iterate through each XmlMap, and output its Name and RootElementName for debugging before saving the file.
// Keywords: Aspose.Cells | C# | XmlMapCollection | enumerate XML maps | RootElementName | XML schema | workbook debugging | .NET | sample code | GitHub example
// Common Searches: Aspose.Cells enumerate XmlMapCollection C# | Get XmlMap Name and RootElementName in .NET | Loop through XML maps Aspose.Cells | Debug XML maps in Excel workbook | Sample code for XmlMapCollection enumeration
// Developer Intent: Retrieve all XML maps in a workbook and display each map's identifier and root element to verify correct mapping.
// Use Cases: Confirm that XML maps were added with the expected names and root elements after workbook creation. | Diagnose schema mismatches by comparing expected root nodes with actual XmlMap.RootElementName values. | Create a diagnostic log of XML map details before exporting or processing the workbook further.
// AI Prompts: Generate C# code that iterates over workbook.Worksheets.XmlMaps and prints each map's Name and RootElementName. | Write a helper method returning a Dictionary<string,string> where the key is XmlMap.Name and the value is XmlMap.RootElementName using Aspose.Cells. | Show how to log XML map details to a file instead of the console for later analysis.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, add an XML schema as an XmlMap, retrieve the workbook's XmlMapCollection, iterate through each XmlMap, and output its Name and RootElementName for debugging before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Define an XML schema as a string.
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write the schema to a temporary file because Aspose.Cells expects a file path.
            string tempSchemaPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
            File.WriteAllText(tempSchemaPath, xmlSchema);

            // Ensure the temporary schema file exists before adding.
            if (File.Exists(tempSchemaPath))
            {
                // Add the XML map to the workbook.
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
                workbook.Worksheets.XmlMaps[mapIndex].Name = "SampleMap";
            }
            else
            {
                Console.WriteLine("Failed to create temporary XML schema file.");
                return;
            }

            // Get the XmlMapCollection from the workbook.
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Enumerate each XmlMap and log its Name and RootElementName.
            for (int i = 0; i < xmlMaps.Count; i++)
            {
                XmlMap map = xmlMaps[i];
                Console.WriteLine($"Map #{i}: Name = {map.Name}, Root Element = {map.RootElementName}");
            }

            // Save the workbook.
            string outputPath = "EnumeratedXmlMaps.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
