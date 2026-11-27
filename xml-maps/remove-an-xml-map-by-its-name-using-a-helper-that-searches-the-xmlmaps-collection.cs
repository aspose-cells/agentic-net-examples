// Title: How to remove an XML map by name from an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# remove XML map by its name from workbook | Delete specific XmlMap from Excel file using Aspose.Cells .NET | Find and remove XML map in Aspose.Cells workbook programmatically | C# code to delete an XML schema map from a workbook with Aspose.Cells | Remove XmlMap entry by name in Aspose.Cells for .xlsx files
// Tags: remove xml map Aspose.Cells C# | xmlmap collection delete Aspose.Cells | excel workbook xml map removal .NET | aspose.cells xml schema map management | c# delete xml map from workbook

using System;
using System.IO;
using Aspose.Cells;

namespace RemoveXmlMapDemo
{
    // The example creates a workbook, adds an XML map from a temporary XSD file, assigns a custom name, then uses a helper method that searches the XmlMaps collection for that name and removes the map before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the workbook that will be saved (new workbook)
                string workbookPath = "SampleWorkbook.xlsx";

                // Create a new workbook
                Workbook workbook = new Workbook();

                // ------------------------------------------------------------
                // Step 1: Add a sample XML map to the workbook
                // ------------------------------------------------------------
                string xmlSchema = @"
                    <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                      <xs:element name='Employees'>
                        <xs:complexType>
                          <xs:sequence>
                            <xs:element name='Employee' maxOccurs='unbounded'>
                              <xs:complexType>
                                <xs:sequence>
                                  <xs:element name='ID' type='xs:int'/>
                                  <xs:element name='Name' type='xs:string'/>
                                </xs:sequence>
                              </xs:complexType>
                            </xs:element>
                          </xs:sequence>
                        </xs:complexType>
                      </xs:element>
                    </xs:schema>";

                // Write the schema to a temporary file because Add(string) expects a file path
                string tempSchemaPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
                File.WriteAllText(tempSchemaPath, xmlSchema);

                // Add the XML map using the temporary file path
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Assign a custom name to the map (this is the name we will later remove)
                string targetMapName = "EmployeeMap";
                xmlMap.Name = targetMapName;

                Console.WriteLine($"Added XML map with name: {xmlMap.Name}");

                // ------------------------------------------------------------
                // Step 2: Remove the XML map by its name using a helper method
                // ------------------------------------------------------------
                bool removed = RemoveXmlMapByName(workbook, targetMapName);
                Console.WriteLine(removed
                    ? $"Successfully removed XML map '{targetMapName}'."
                    : $"XML map '{targetMapName}' was not found.");

                // ------------------------------------------------------------
                // Step 3: Save the workbook (optional, just to verify the map is gone)
                // ------------------------------------------------------------
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <param name="workbook">The workbook containing the XmlMaps collection.</param>
        /// <param name="mapName">The name of the XML map to remove.</param>
        /// <returns>True if the map was found and removed; otherwise false.</returns>
        private static bool RemoveXmlMapByName(Workbook workbook, string mapName)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrEmpty(mapName)) throw new ArgumentException("Map name cannot be null or empty.", nameof(mapName));

            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;
            int targetIndex = -1;

            // Linear search for the map name (case‑sensitive)
            for (int i = 0; i < xmlMaps.Count; i++)
            {
                if (xmlMaps[i].Name == mapName)
                {
                    targetIndex = i;
                    break;
                }
            }

            if (targetIndex >= 0)
            {
                xmlMaps.RemoveAt(targetIndex);
                return true;
            }

            // Map not found
            return false;
        }
    }
}
