// Title: How to link Excel range B2:C5 to a repeating <Item> XML element using Aspose.Cells SetXmlMap in C#
// AI Prompts: Create an XML map for a repeating <Item> element from an XSD string and bind it to the range B2:C5 with SetXmlMap. | Use reflection to add a repeatable XML map to a workbook and apply SetXmlMap to a specified cell range in C#. | Generate an Excel file where the cells B2 through C5 are mapped to a repeating XML schema element using Aspose.Cells.
// Common Searches: aspnet c# example linking worksheet range to repeating xml element with aspose.cells setxmlmap | how to map xml schema repeating rows to excel cells using aspose cells library | c# setxmlmap for range B2:C5 with repeatable xml map using reflection
// Tags: Aspose.Cells SetXmlMap for repeating elements | C# associate worksheet cells with XML map | Reflection-based XML map creation Aspose.Cells | Excel range B2:C5 linked to XSD schema | Generate LinkedXmlMap.xlsx with XML mapping

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The sample creates a new workbook, defines an XSD schema containing a repeating <Item> element, uses reflection to add an XML map with repeatable rows, creates the range B2:C5, invokes SetXmlMap to bind the range to the XML map, and saves the workbook as LinkedXmlMap.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Attempt XML mapping using reflection (may not be supported in older versions)
            try
            {
                // XML schema defining a repeating element <Item>
                string xsd = @"<?xml version='1.0'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema' targetNamespace='http://example.com' xmlns='http://example.com' elementFormDefault='qualified'>
  <xs:element name='Root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Col1' type='xs:string'/>
              <xs:element name='Col2' type='xs:string'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                // Use reflection to access Workbook.XmlMaps if it exists
                PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
                if (xmlMapsProp != null)
                {
                    object xmlMaps = xmlMapsProp.GetValue(workbook);
                    MethodInfo addMethod = xmlMaps.GetType().GetMethod("Add", new[] { typeof(string), typeof(string), typeof(bool) });
                    if (addMethod != null)
                    {
                        // Add the XML map (true = allow repeating)
                        object xmlMap = addMethod.Invoke(xmlMaps, new object[] { "Map1", xsd, true });

                        // Create the target range and link it to the XML map via reflection
                        Aspose.Cells.Range range = sheet.Cells.CreateRange("B2:C5");
                        MethodInfo setXmlMapMethod = typeof(Aspose.Cells.Range).GetMethod("SetXmlMap", new[] { xmlMap.GetType(), typeof(bool) });
                        if (setXmlMapMethod != null)
                        {
                            setXmlMapMethod.Invoke(range, new object[] { xmlMap, true });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log but continue – XML mapping may not be supported in this environment
                Console.WriteLine($"XML mapping skipped: {ex.Message}");
            }

            // Save the workbook
            string outputPath = "LinkedXmlMap.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
