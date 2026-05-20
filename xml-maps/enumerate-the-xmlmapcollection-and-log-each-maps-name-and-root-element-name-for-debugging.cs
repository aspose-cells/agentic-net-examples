using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample XML schema (replace with actual schema or file path as needed)
                string sampleSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
    <xs:element name='Root'>
        <xs:complexType>
            <xs:sequence>
                <xs:element name='Item' type='xs:string'/>
            </xs:sequence>
        </xs:complexType>
    </xs:element>
</xs:schema>";

                // Add the XML map from the schema string
                int mapIndex = workbook.Worksheets.XmlMaps.Add(sampleSchema);
                XmlMap addedMap = workbook.Worksheets.XmlMaps[mapIndex];
                addedMap.Name = "SampleMap";

                // Enumerate all XmlMaps in the workbook
                for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
                {
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[i];
                    Console.WriteLine($"XmlMap #{i + 1}");
                    Console.WriteLine($"  Name: {xmlMap.Name}");
                    Console.WriteLine($"  Root Element Name: {xmlMap.RootElementName}");
                }

                // Save the workbook
                string outputPath = "XmlMapEnumerationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}