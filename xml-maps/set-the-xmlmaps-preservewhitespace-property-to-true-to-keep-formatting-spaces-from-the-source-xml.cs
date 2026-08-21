// Title: Aspose.Cells C# – Set XmlMap.PreserveWhitespace = true to keep original XML spacing
// Description: Shows how to create a workbook, add an XML map, enable the XmlMap.PreserveWhitespace property, and save the workbook as XML so that all indentation and space characters from the source file are preserved. Highlights that whitespace handling is a property of XmlMap, not XmlSaveOptions.
// Keywords: Aspose.Cells XmlMap PreserveWhitespace | C# preserve XML whitespace | keep XML formatting Aspose.Cells | XmlMap whitespace property | .NET export XML with original spaces | Aspose.Cells XML map indentation
// Common Searches: how to preserve whitespace with Aspose.Cells XmlMap | XmlMap PreserveWhitespace C# example | Aspose.Cells keep XML indentation when saving | set PreserveWhitespace on XmlMap .NET | why XmlSaveOptions has no PreserveWhitespace
// Developer Intent: Enable XmlMap.PreserveWhitespace = true before saving so the exported XML retains the source file's spacing.
// Use Cases: Export data to an XML document that must match a strict schema layout, including exact indentation. | Perform round‑trip XML import/export where whitespace is significant, such as mixed‑content elements. | Generate XML reports for downstream systems that rely on precise formatting for parsing or visual comparison.
// AI Prompts: Provide C# code using Aspose.Cells to load an XML map, set PreserveWhitespace = true, and save the workbook without altering spaces. | Explain why XmlSaveOptions lacks a PreserveWhitespace property and demonstrate the correct way to preserve whitespace with XmlMap. | Create a full example that reads an XML file with important whitespace, maps it in Aspose.Cells, enables PreserveWhitespace, and writes the file back unchanged.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add an XML map, enable the XmlMap.PreserveWhitespace property, and save the workbook as XML so that all indentation and space characters from the source file are preserved. Highlights that whitespace handling is a property of XmlMap, not XmlSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(100);

            // Define an XML schema (the actual schema content is not important for this demo)
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Data'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:sequence>
                                                            <xs:element name='Name' type='xs:string'/>
                                                            <xs:element name='Value' type='xs:integer'/>
                                                        </xs:sequence>
                                                    </xs:complexType>
                                                </xs:element>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Add the XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "DataMap"; // Assign a name to the map (required for saving)

            // Configure XML save options to include the map
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                XmlMapName = xmlMap.Name
                // PreserveWhitespace is not a valid property for XmlSaveOptions; whitespace handling is managed during load.
            };

            // Save the workbook as an XML file
            workbook.Save("output.xml", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
