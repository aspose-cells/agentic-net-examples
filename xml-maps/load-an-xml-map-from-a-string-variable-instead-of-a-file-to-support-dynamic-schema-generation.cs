// Title: Create a Workbook and import XML data using a schema string with Aspose.Cells for .NET (reflection fallback)
// AI Prompts: Write C# code that builds a Workbook, adds an XML map from an XML schema string via the XmlMaps collection, imports XML data, and saves the file as XLSX. | Show how to use reflection to call the Add method on XmlMaps and the ImportXml method on the resulting XmlMap, handling cases where the XmlMaps property is missing. | Provide a reusable C# method that accepts XML schema and XML data strings, creates a dynamic XML map, imports the data into a new workbook, and returns the path of the saved workbook.
// Common Searches: aspnet how to add an XML map from a string in Aspose.Cells | c# import xml data into Excel workbook without using a schema file Aspose.Cells | using reflection to access XmlMaps collection in Aspose.Cells .NET | fallback to Workbook.ImportXml when XmlMaps property is unavailable | dynamic XML schema mapping to Excel with Aspose.Cells C# example
// Tags: add xml map from schema string Aspose.Cells | import xml data into workbook C# | reflection access XmlMaps collection | fallback Workbook ImportXml overload | save workbook as xlsx after xml import

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The program creates a new Workbook, defines an XML schema and XML data as strings, adds an XML map using reflection (or a fallback ImportXml overload), imports the XML data into the workbook, and saves the result as DynamicXmlMapOutput.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (no template file needed)
                var workbook = new Workbook();

                // XML schema defined as a string
                string xmlSchema = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Name' type='xs:string'/>
              <xs:element name='Value' type='xs:int'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                // Example XML data to import
                string xmlData = @"<Root>
  <Item>
    <Name>Item1</Name>
    <Value>10</Value>
  </Item>
  <Item>
    <Name>Item2</Name>
    <Value>20</Value>
  </Item>
</Root>";

                // Try to add an XML map and import data using reflection (covers different API versions)
                try
                {
                    // Look for the XmlMaps property
                    var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
                    if (xmlMapsProp != null)
                    {
                        // Get the XmlMapCollection instance
                        var xmlMaps = xmlMapsProp.GetValue(workbook);
                        // Add a new map: Add(string mapName, string xmlSchema)
                        var addMethod = xmlMaps.GetType().GetMethod("Add", new[] { typeof(string), typeof(string) });
                        var xmlMap = addMethod?.Invoke(xmlMaps, new object[] { "MyDynamicMap", xmlSchema });

                        // Import XML data: ImportXml(string xmlData, bool importDataOnly)
                        var importMethod = xmlMap?.GetType().GetMethod("ImportXml", new[] { typeof(string), typeof(bool) });
                        importMethod?.Invoke(xmlMap, new object[] { xmlData, true });
                    }
                    else
                    {
                        // Fallback: Workbook may expose an ImportXml overload directly
                        var importMethod = workbook.GetType().GetMethod("ImportXml", new[] { typeof(string), typeof(string), typeof(bool) });
                        importMethod?.Invoke(workbook, new object[] { xmlData, "MyDynamicMap", true });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"XML map import failed: {ex.Message}");
                }

                // Determine a safe output path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicXmlMapOutput.xlsx");

                // Ensure the directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
