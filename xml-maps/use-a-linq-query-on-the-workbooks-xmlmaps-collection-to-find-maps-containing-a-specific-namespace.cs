using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsXmlMapQuery
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare dummy XSD files if they do not exist to avoid FileNotFoundException
                string schemaPath1 = "schema1.xsd";
                string schemaPath2 = "schema2.xsd";

                if (!File.Exists(schemaPath1))
                {
                    File.WriteAllText(schemaPath1,
                        @"<?xml version=""1.0"" encoding=""utf-8""?>
                        <xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema""></xs:schema>");
                }

                if (!File.Exists(schemaPath2))
                {
                    File.WriteAllText(schemaPath2,
                        @"<?xml version=""1.0"" encoding=""utf-8""?>
                        <xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema""></xs:schema>");
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add XML maps with different namespaces (using local XSD files)
                int mapIndex1 = workbook.Worksheets.XmlMaps.Add(schemaPath1);
                XmlMap xmlMap1 = workbook.Worksheets.XmlMaps[mapIndex1];
                xmlMap1.Name = "Map1";

                int mapIndex2 = workbook.Worksheets.XmlMaps.Add(schemaPath2);
                XmlMap xmlMap2 = workbook.Worksheets.XmlMaps[mapIndex2];
                xmlMap2.Name = "Map2";

                // Define the namespace (or part of it) to search for
                string targetNamespace = "schema1.xsd";

                // LINQ query on the XmlMaps collection to find maps containing the specific namespace
                List<XmlMap> matchingMaps = workbook.Worksheets.XmlMaps
                    .Cast<XmlMap>()
                    .Where(m => m.DataBinding != null &&
                                !string.IsNullOrEmpty(m.DataBinding.Url) &&
                                m.DataBinding.Url.Contains(targetNamespace, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                // Output the results
                foreach (XmlMap map in matchingMaps)
                {
                    Console.WriteLine($"Found map: Name = {map.Name}, URL = {map.DataBinding.Url}");
                }

                // Save the workbook (output file)
                workbook.Save("XmlMapNamespaceQuery.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}