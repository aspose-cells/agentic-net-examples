// Title: Import XML data into an Excel workbook using an XSD schema map with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# import XML file using XSD map into worksheet | How to add an XML map from an XSD schema and load XML data with Workbook.ImportXml | C# example for mapping XSD to Excel and importing XML using Aspose.Cells | Using Workbook.Worksheets.XmlMaps.Add to import XML into Sheet1 with Aspose.Cells | Save imported XML data to XLSX after adding XSD map in C#
// Tags: xsd schema to xml map Aspose.Cells | Workbook.ImportXml with predefined xml map | excel worksheet schema mapping c# | populate sheet from xml using aspocells | export workbook to xlsx after xml import

using System;
using System.IO;
using Aspose.Cells;

// The example creates temporary XSD and XML files, adds the XSD as an XML map to a new Workbook via Worksheets.XmlMaps.Add, imports the XML data into Sheet1 starting at cell A1 using Workbook.ImportXml, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Prepare sample XSD (XML schema) that defines the structure of the XML data.
        string xsdContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Products'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Product' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Name' type='xs:string'/>
              <xs:element name='Price' type='xs:decimal'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        // Prepare sample XML that conforms to the above XSD.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Products>
  <Product>
    <Name>Laptop</Name>
    <Price>999.99</Price>
  </Product>
  <Product>
    <Name>Phone</Name>
    <Price>699.99</Price>
  </Product>
</Products>";

        // Define temporary file paths (they will be created in the same folder as the executable).
        string xsdPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.xsd");
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.xml");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ImportedFromXsd.xlsx");

        try
        {
            // Write the XSD and XML strings to physical files.
            File.WriteAllText(xsdPath, xsdContent);
            File.WriteAllText(xmlPath, xmlContent);

            // Create a new workbook (initially empty).
            Workbook workbook = new Workbook();

            // ------------------------------------------------------------
            // Step 1: Add the XSD as an XML map to the workbook.
            // ------------------------------------------------------------
            // The XmlMaps collection belongs to the Worksheets collection.
            // Adding the XSD returns the index of the newly created map.
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xsdPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            // Optionally give the map a friendly name.
            xmlMap.Name = "ProductsMap";

            Console.WriteLine($"XML map '{xmlMap.Name}' added. Root element: {xmlMap.RootElementName}");

            // ------------------------------------------------------------
            // Step 2: Import the XML data using the map.
            // ------------------------------------------------------------
            // ImportXml will create the necessary rows/columns based on the XML structure.
            // Parameters: (xmlFilePath, destinationSheetName, startRow, startColumn)
            workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

            Console.WriteLine("XML data imported successfully into Sheet1 starting at cell A1.");

            // ------------------------------------------------------------
            // Step 3: Save the workbook to verify the result.
            // ------------------------------------------------------------
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        finally
        {
            // Clean up temporary files (optional).
            try { if (File.Exists(xsdPath)) File.Delete(xsdPath); } catch { }
            try { if (File.Exists(xmlPath)) File.Delete(xmlPath); } catch { }
        }
    }
}
