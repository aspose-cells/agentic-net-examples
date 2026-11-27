// Title: Import XML data into linked Excel cells with Workbook.ImportXml after adding an XML map in C# (Aspose.Cells)
// Common Searches: C# Aspose.Cells import XML file into worksheet after creating XML map from XSD | Using Workbook.ImportXml to bind XML data to worksheet cells in Aspose.Cells .NET | Add XML map to a workbook and import XML starting at cell A1 with Aspose.Cells | How to link XML elements to Excel cells using Aspose.Cells ImportXml method | Sample code for importing XML into Excel using Aspose.Cells and XmlMap
// Tags: Workbook.ImportXml with pre‑added XmlMap | XML map creation from XSD Aspose.Cells | import XML into Excel worksheet C# | link XML elements to cells Aspose.Cells | Aspose.Cells XML import example

using System;
using System.IO;
using Aspose.Cells;

// The example creates temporary XSD and XML files, adds the XSD as an XML map to a new Workbook, then uses Workbook.ImportXml to import the XML data into the first worksheet starting at cell A1, demonstrating linked cells by reading values, saves the workbook as an XLSX file, and cleans up the temporary files.
class Program
{
    static void Main()
    {
        // Paths for temporary schema and data files
        string schemaPath = Path.Combine(Path.GetTempPath(), "sample.xsd");
        string xmlPath = Path.Combine(Path.GetTempPath(), "sample.xml");
        string outputPath = "ImportedXmlWorkbook.xlsx";

        // Sample XML schema (XSD) that defines the structure of the XML data
        string xmlSchema = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Products'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Product' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Name' type='xs:string' />
              <xs:element name='Price' type='xs:decimal' />
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        // Sample XML data that conforms to the above schema
        string xmlData = @"<?xml version='1.0' encoding='utf-8'?>
<Products>
  <Product>
    <Name>Laptop</Name>
    <Price>999.99</Price>
  </Product>
  <Product>
    <Name>Phone</Name>
    <Price>699.99</Price>
  </Product>
  <Product>
    <Name>Tablet</Name>
    <Price>399.99</Price>
  </Product>
</Products>";

        try
        {
            // Write schema and XML data to temporary files
            File.WriteAllText(schemaPath, xmlSchema);
            File.WriteAllText(xmlPath, xmlData);

            // Create a new workbook (initially empty)
            Workbook workbook = new Workbook();

            // Add the XML map to the workbook using the XSD file
            int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            Console.WriteLine($"XML Map added. Name: {xmlMap.Name}, Root Element: {xmlMap.RootElementName}");

            // Import the XML data into the first worksheet starting at cell A1 (row 0, column 0)
            // The ImportXml method will use the previously added XML map to bind data to cells
            workbook.ImportXml(xmlPath, "Sheet1", 0, 0);
            Console.WriteLine("XML data imported into worksheet.");

            // Optional: demonstrate that the data is linked by retrieving a cell value
            Worksheet sheet = workbook.Worksheets["Sheet1"];
            Console.WriteLine($"Cell A2 (first product name): {sheet.Cells["A2"].StringValue}");
            Console.WriteLine($"Cell B2 (first product price): {sheet.Cells["B2"].StringValue}");

            // Save the workbook to the specified output file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up temporary files
            try { if (File.Exists(schemaPath)) File.Delete(schemaPath); } catch { }
            try { if (File.Exists(xmlPath)) File.Delete(xmlPath); } catch { }
        }
    }
}
