// Title: Load an Excel workbook from a MemoryStream, attach an XSD‑based XML map, and save it back to a MemoryStream using Aspose.Cells for .NET
// Common Searches: Aspose.Cells add XML map to workbook loaded from MemoryStream C# | save Excel file with XML map to MemoryStream using Aspose.Cells | load workbook from stream and associate XSD schema as XML map Aspose.Cells | in‑memory Excel processing add XML map from XSD Aspose.Cells .NET
// Tags: memory stream workbook loading Aspose.Cells | xml map addition from XSD Aspose.Cells | save workbook to memory stream C# | in‑memory Excel manipulation with XML schema | xml maps collection add method Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    // Demonstrates how to create an Excel workbook in memory, load it from a MemoryStream, write an XSD schema to a temporary file, add the schema as an XML map via the XmlMaps collection, and then save the updated workbook back to a new MemoryStream (optionally writing it to disk) using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Ensure Aspose.Cells license is set if you have one.
            // License license = new License();
            // license.SetLicense("Aspose.Cells.lic");

            try
            {
                // ------------------------------------------------------------
                // 1. Create a sample workbook in memory (this simulates an existing file)
                // ------------------------------------------------------------
                Workbook sourceWorkbook = new Workbook();
                Worksheet srcSheet = sourceWorkbook.Worksheets[0];
                srcSheet.Name = "DataSheet";

                // Populate some sample data
                srcSheet.Cells["A1"].PutValue("Name");
                srcSheet.Cells["B1"].PutValue("Value");
                srcSheet.Cells["A2"].PutValue("Item1");
                srcSheet.Cells["B2"].PutValue(100);
                srcSheet.Cells["A3"].PutValue("Item2");
                srcSheet.Cells["B3"].PutValue(200);

                // Save the source workbook into a memory stream
                using (MemoryStream sourceStream = new MemoryStream())
                {
                    sourceWorkbook.Save(sourceStream, SaveFormat.Xlsx);
                    sourceStream.Position = 0; // Reset for reading

                    // ------------------------------------------------------------
                    // 2. Load the workbook from the memory stream
                    // ------------------------------------------------------------
                    Workbook workbook = new Workbook(sourceStream);

                    // ------------------------------------------------------------
                    // 3. Prepare an XML schema (XSD) that will be used as an XML map
                    // ------------------------------------------------------------
                    string xsdContent = @"
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Data'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Name' type='xs:string'/>
              <xs:element name='Value' type='xs:integer'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                    // Write the XSD to a temporary file because XmlMapCollection.Add expects a file path
                    string tempXsdPath = Path.Combine(Path.GetTempPath(), "SampleMap.xsd");
                    File.WriteAllText(tempXsdPath, xsdContent);

                    // ------------------------------------------------------------
                    // 4. Add the XML map to the workbook
                    // ------------------------------------------------------------
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                    xmlMap.Name = "SampleMap";

                    // Clean up the temporary XSD file
                    if (File.Exists(tempXsdPath))
                    {
                        File.Delete(tempXsdPath);
                    }

                    // ------------------------------------------------------------
                    // 5. Save the modified workbook back to a new memory stream
                    // ------------------------------------------------------------
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        workbook.Save(outputStream, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook with XML map saved to memory stream. Stream length: {outputStream.Length} bytes.");

                        // (Optional) Write the stream to a physical file for verification
                        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WorkbookWithXmlMap.xlsx");
                        File.WriteAllBytes(outputPath, outputStream.ToArray());
                        Console.WriteLine($"Workbook written to file: {outputPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
