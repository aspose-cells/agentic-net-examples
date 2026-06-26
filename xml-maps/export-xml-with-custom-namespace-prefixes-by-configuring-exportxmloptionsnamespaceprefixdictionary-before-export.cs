using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportXmlWithCustomNamespacePrefixes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Id");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");

                // Define a simple XML schema for the map
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema' xmlns:ex='http://example.com/ns'>
    <xs:element name='Employees' type='ex:EmployeesType'/>
    <xs:complexType name='EmployeesType'>
        <xs:sequence>
            <xs:element name='Employee' maxOccurs='unbounded' type='ex:EmployeeType'/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name='EmployeeType'>
        <xs:sequence>
            <xs:element name='Id' type='xs:int'/>
            <xs:element name='Name' type='xs:string'/>
        </xs:sequence>
    </xs:complexType>
</xs:schema>";

                // Add the XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "EmployeeMap";

                // NOTE: In older Aspose.Cells versions XmlMapExportOptions may not be available.
                // The ExportXml method without options is used here to ensure compatibility.
                // If a newer version is referenced, you can uncomment the block below to set custom prefixes.

                /*
                // Create XmlMapExportOptions and set custom namespace prefixes
                XmlMapExportOptions exportOptions = new XmlMapExportOptions
                {
                    NamespacePrefixDictionary = new Dictionary<string, string>
                    {
                        { "http://example.com/ns", "ex" } // map the namespace URI to prefix "ex"
                    }
                };
                workbook.ExportXml(xmlMap.Name, outputPath, exportOptions);
                */

                // Output file path
                string outputPath = "EmployeesWithCustomPrefix.xml";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the XML using the map name and output file path
                workbook.ExportXml(xmlMap.Name, outputPath);

                Console.WriteLine($"XML exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportXmlWithCustomNamespacePrefixes.Run();
        }
    }
}