// Title: Create a named XML map in an Excel workbook with Aspose.Cells C# using XmlMapCollection.Add
// Common Searches: how to add an XML map with a custom name to a workbook using Aspose.Cells C# | Aspose.Cells C# XmlMapCollection.Add example for assigning map name | set custom name for an XML map after adding XSD schema in Aspose.Cells | C# code to create and save an Excel file with a named XML map using Aspose.Cells | adding XSD as XML map and renaming it in Aspose.Cells workbook
// Tags: Aspose.Cells XmlMapCollection Add C# | custom XML map name Aspose.Cells | add XSD schema as XML map C# | save workbook with XML map Aspose.Cells | populate worksheet data matching XML schema Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a temporary XSD schema, initializes a new Workbook, adds the XSD as an XmlMap via XmlMapCollection.Add, assigns a custom name to the XmlMap, optionally fills worksheet cells that conform to the schema, saves the workbook as an .xlsx file, and cleans up the temporary XSD file.
class Program
{
    static void Main()
    {
        // Paths for temporary files and final workbook
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeCellsXmlMapDemo");
        Directory.CreateDirectory(tempFolder);

        string xsdPath = Path.Combine(tempFolder, "sample.xsd");
        string workbookPath = Path.Combine(tempFolder, "WorkbookWithXmlMap.xlsx");

        try
        {
            // 1. Create a simple XSD schema that describes a very small XML structure.
            string xsdContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""Root"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""Item"" maxOccurs=""unbounded"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""Name"" type=""xs:string""/>
              <xs:element name=""Value"" type=""xs:int""/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

            // Write XSD to temporary file
            File.WriteAllText(xsdPath, xsdContent);
            Console.WriteLine($"XSD schema written to: {xsdPath}");

            // 2. Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // 3. Access the XmlMapCollection from the workbook
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // 4. Add the XSD file as a new XmlMap. The Add method returns the index of the newly added map.
            if (!File.Exists(xsdPath))
                throw new FileNotFoundException("XSD file not found.", xsdPath);

            int mapIndex = xmlMaps.Add(xsdPath);
            Console.WriteLine($"XmlMap added at index {mapIndex}");

            // 5. Retrieve the XmlMap object and assign a custom name.
            XmlMap xmlMap = xmlMaps[mapIndex];
            string customMapName = "MyCustomMap";
            xmlMap.Name = customMapName;
            Console.WriteLine($"Custom name set for XmlMap: {xmlMap.Name}");

            // 6. (Optional) Populate some data in the first worksheet that matches the schema.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Item2");
            sheet.Cells["B3"].PutValue(200);

            // 7. Save the workbook so that the XmlMap definition is persisted.
            workbook.Save(workbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {workbookPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
        finally
        {
            // Clean up the temporary XSD file (keep the workbook for inspection)
            try
            {
                if (File.Exists(xsdPath))
                {
                    File.Delete(xsdPath);
                    Console.WriteLine($"Temporary XSD file deleted: {xsdPath}");
                }
            }
            catch (Exception cleanupEx)
            {
                Console.WriteLine("Failed to delete temporary XSD file:");
                Console.WriteLine(cleanupEx.Message);
            }
        }

        Console.WriteLine("Demo completed. Press any key to exit.");
        Console.ReadKey();
    }
}
