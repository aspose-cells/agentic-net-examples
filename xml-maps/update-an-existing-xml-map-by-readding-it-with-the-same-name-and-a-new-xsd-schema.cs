using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapUpdate
{
    class Program
    {
        static void Main()
        {
            // Paths
            string workbookPath = "InputWorkbook.xlsx";
            string outputPath = "UpdatedWorkbook.xlsx";

            // Name of the XML map to replace
            string mapNameToUpdate = "MyXmlMap";

            // New XSD schema as a string
            string newXsdSchema = @"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""Root"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""NewElement"" type=""xs:string""/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

            try
            {
                // Verify the input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access XML maps collection
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Find existing map index
                int existingMapIndex = -1;
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    if (xmlMaps[i].Name == mapNameToUpdate)
                    {
                        existingMapIndex = i;
                        break;
                    }
                }

                // Remove existing map if found
                if (existingMapIndex != -1)
                {
                    xmlMaps.RemoveAt(existingMapIndex);
                }

                // Write the XSD schema to a temporary file (required by Aspose.Cells API)
                string tempXsdPath = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.xsd");
                File.WriteAllText(tempXsdPath, newXsdSchema);

                // Add the new XML map using the temporary XSD file
                int newMapIndex = xmlMaps.Add(tempXsdPath);

                // Set the desired name for the new map
                XmlMap newMap = xmlMaps[newMapIndex];
                newMap.Name = mapNameToUpdate;

                // Clean up temporary XSD file
                if (File.Exists(tempXsdPath))
                {
                    File.Delete(tempXsdPath);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}