using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapDebug
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the XmlMapCollection from the workbook
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Define simple XSD schemas as strings
                string schema1 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                      <xs:element name='Root1'/>
                                   </xs:schema>";
                string schema2 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                      <xs:element name='Root2'/>
                                   </xs:schema>";

                // Write schemas to temporary files (Add expects a file path)
                string tempFile1 = Path.GetTempFileName();
                string tempFile2 = Path.GetTempFileName();
                File.WriteAllText(tempFile1, schema1);
                File.WriteAllText(tempFile2, schema2);

                // Ensure the temporary files exist before adding
                if (File.Exists(tempFile1) && File.Exists(tempFile2))
                {
                    int mapIndex1 = xmlMaps.Add(tempFile1);
                    int mapIndex2 = xmlMaps.Add(tempFile2);

                    // Set friendly names for the maps
                    xmlMaps[mapIndex1].Name = "FirstMap";
                    xmlMaps[mapIndex2].Name = "SecondMap";

                    // Enumerate the XmlMapCollection and display each map's details
                    for (int i = 0; i < xmlMaps.Count; i++)
                    {
                        XmlMap map = xmlMaps[i];
                        Console.WriteLine($"Map {i}: Name = {map.Name}, Root Element = {map.RootElementName}");
                    }

                    // Save the workbook
                    string outputPath = "XmlMapDebugOutput.xlsx";
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
                else
                {
                    Console.WriteLine("Failed to create temporary schema files.");
                }

                // Clean up temporary files
                if (File.Exists(tempFile1)) File.Delete(tempFile1);
                if (File.Exists(tempFile2)) File.Delete(tempFile2);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}