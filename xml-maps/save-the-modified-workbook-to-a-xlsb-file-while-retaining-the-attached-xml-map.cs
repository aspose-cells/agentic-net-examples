// Title: Save a Workbook as XLSB with Embedded XML Map using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, attach an XML map from an XSD schema, and export it to a binary .xlsb file with XlsbSaveOptions while preserving the XML mapping metadata.
// Keywords: Aspose.Cells | C# | .NET | XLSB export | XML map | XSD schema | XlsbSaveOptions | preserve XML map | binary Excel file | save workbook as .xlsb
// Common Searches: Aspose.Cells save workbook as xlsb with xml map | C# keep xml map when exporting to xlsb | how to embed xml schema in xlsb using Aspose.Cells | XlsbSaveOptions retain xml mapping | export Excel to binary format with xml map .NET
// Developer Intent: Export a workbook to .xlsb while ensuring the attached XML map remains intact.
// Use Cases: Distribute a binary Excel template that includes XML mapping for automated data import. | Generate reports for systems that require XLSB files with embedded XML schema metadata. | Create reusable XLSB files for downstream applications that consume XML‑mapped data.
// AI Prompts: Write C# code with Aspose.Cells to add an XML map from an XSD file and save the workbook as XLSB, keeping the map embedded. | Explain the role of XlsbSaveOptions in preserving XML maps during XLSB export and list any required properties. | Show how to programmatically verify that an XML map exists in a saved XLSB file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, attach an XML map from an XSD schema, and export it to a binary .xlsb file with XlsbSaveOptions while preserving the XML mapping metadata.
class SaveXlsbWithXmlMap
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(999.99);

            // Define a simple XML schema (XSD) and write it to a temporary file
            string xmlSchema = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Product'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Name' type='xs:string' />
        <xs:element name='Price' type='xs:decimal' />
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

            string schemaPath = "ProductSchema.xsd";

            // Ensure the schema file exists (create or overwrite)
            File.WriteAllText(schemaPath, xmlSchema);

            // Add the XML map using the schema file
            int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "ProductDataMap"; // give the map a friendly name

            // Create XLSB save options
            XlsbSaveOptions saveOptions = new XlsbSaveOptions();

            // Save the workbook as an XLSB file while retaining the attached XML map
            string outputPath = "ProductData.xlsb";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
