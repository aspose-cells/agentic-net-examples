using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        XmlMapPrerequisiteCheck.Run();
    }
}

class XmlMapPrerequisiteCheck
{
    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the XmlMap collection from the workbook
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        // Verify that the collection is available
        if (xmlMaps == null)
        {
            throw new InvalidOperationException("XmlMap collection is not available.");
        }

        // Sample XML schema that will be used to create the map
        string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                <xs:element name='Root'>
                                    <xs:complexType>
                                        <xs:sequence>
                                            <xs:element name='Item' type='xs:string'/>
                                        </xs:sequence>
                                    </xs:complexType>
                                </xs:element>
                            </xs:schema>";

        // Write the schema to a temporary file because Add expects a file path
        string tempXsdPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
        File.WriteAllText(tempXsdPath, xmlSchema);

        // Check whether a map with the intended name already exists
        const string desiredMapName = "RootMap";
        bool mapExists = false;
        foreach (XmlMap existingMap in xmlMaps)
        {
            if (existingMap.Name == desiredMapName)
            {
                mapExists = true;
                break;
            }
        }

        if (!mapExists)
        {
            // Add the XML map to the collection using the temporary XSD file
            int mapIndex = xmlMaps.Add(tempXsdPath);
            XmlMap newMap = xmlMaps[mapIndex];
            newMap.Name = desiredMapName;
            Console.WriteLine("XML map added successfully.");
        }
        else
        {
            Console.WriteLine("XML map already exists; no action required.");
        }

        // Clean up the temporary XSD file
        if (File.Exists(tempXsdPath))
        {
            File.Delete(tempXsdPath);
        }

        // Optional: save the workbook to verify that everything works
        workbook.Save("PrereqCheck.xlsx", SaveFormat.Xlsx);
    }
}