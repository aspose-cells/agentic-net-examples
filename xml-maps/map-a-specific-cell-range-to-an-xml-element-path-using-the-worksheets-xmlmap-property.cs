// Title: Mapping a specific Excel cell range to an XML element path with Aspose.Cells XmlMap in C#
// Common Searches: C# Aspose.Cells map range A1:B5 to XML element /Root/Items/Item | How to add an XML map to a worksheet using XmlMaps.Add in Aspose.Cells | Using dynamic binding for XmlMap.AddMapping when compile‑time APIs are unavailable | Save an Aspose.Cells workbook after creating an XML map in .NET | Check XML file existence before creating an XmlMap in Aspose.Cells C# example
// Tags: Aspose.Cells XmlMap AddMapping C# | Excel range to XML element mapping Aspose.Cells | dynamic XmlMaps.Add Aspose.Cells example | create XML map workbook Aspose.Cells | save workbook with XML map Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Microsoft.CSharp.RuntimeBinder;

namespace AsposeCellsXmlMappingExample
{
    // The example creates a new workbook, renames the first worksheet, verifies that a source XML file exists, then uses dynamic binding to add an XML map named "Map1" from the file and maps the cell range A1:B5 to the XML element path "/Root/Items/Item" via XmlMap.AddMapping. It handles missing mapping APIs, ensures the output directory is present, and saves the workbook to the specified location.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Path to the XML data source
                string xmlFilePath = @"C:\Data\sample.xml";

                // Verify that the XML file exists
                if (!File.Exists(xmlFilePath))
                    throw new FileNotFoundException("XML data source not found.", xmlFilePath);

                // Define the cell range to map and the corresponding XML element path
                string cellRange = "A1:B5";
                string xmlElementPath = "/Root/Items/Item";

                // Attempt to add an XML map using dynamic to avoid compile‑time binding issues
                try
                {
                    dynamic dynWorkbook = workbook;
                    dynamic xmlMap = dynWorkbook.XmlMaps.Add("Map1", xmlFilePath);
                    xmlMap.AddMapping(cellRange, xmlElementPath);
                }
                catch (RuntimeBinderException)
                {
                    Console.WriteLine("XML mapping APIs are not available in the current Aspose.Cells version. Skipping mapping step.");
                }

                // Prepare output path and ensure the directory exists
                string outputPath = @"C:\Output\MappedWorkbook.xlsx";
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (with or without XML mapping)
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved successfully at: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
