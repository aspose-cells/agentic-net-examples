// Title: Import XML data into an Excel workbook and link worksheet cells to XML map elements with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells import XML file into worksheet using an XML map | How to bind Excel cells to XML schema elements with Aspose.Cells | Retrieve cells linked to a specific XML map query path in Aspose.Cells C# | Save workbook after importing XML stream using Aspose.Cells .NET | GetAreasOfXmlMapQuery example for locating linked cells in Aspose.Cells
// Tags: Aspose.Cells XML map import C# | link worksheet cells to XML schema elements | GetAreasOfXmlMapQuery usage | import XML stream into Excel workbook | save workbook as XLSX after XML import | XML map query path cell retrieval Aspose.Cells | add XML map from XSD string Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Collections;

// The example creates a new workbook, adds an XML map from an XSD string, imports XML data from a memory stream into the first worksheet, uses GetAreasOfXmlMapQuery to find cells linked to the '/Products/Product/Name' path, prints the linked cell values, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Path to the output workbook
        const string outputFile = "XmlMapLinkedWorkbook.xlsx";

        // Sample XML schema (XSD) defining the structure of the XML data
        string xmlSchema = @"
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

        // Sample XML data that conforms to the above schema
        string xmlData = @"
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

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Products";

        try
        {
            // ------------------------------------------------------------
            // 1. Add the XML map to the workbook using the schema string.
            // ------------------------------------------------------------
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            Console.WriteLine($"XML Map added. Name: {xmlMap.Name}, Root element: {xmlMap.RootElementName}");

            // ------------------------------------------------------------
            // 2. Import the XML data into the worksheet starting at cell A1.
            //    ImportXml will use the previously added XmlMap to map elements.
            // ------------------------------------------------------------
            using (MemoryStream xmlStream = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(xmlStream))
            {
                writer.Write(xmlData);
                writer.Flush();
                xmlStream.Position = 0; // Reset stream position for reading

                // Import XML data into the worksheet. The map is automatically applied.
                workbook.ImportXml(xmlStream, worksheet.Name, 0, 0);
            }

            // ------------------------------------------------------------
            // 3. Verify the linking by querying the map for a specific element path.
            //    Use the modern GetAreasOfXmlMapQuery method.
            // ------------------------------------------------------------
            // Path to the <Name> element inside each <Product>
            string queryPath = "/Products/Product/Name";

            CellArea[] linkedAreas = worksheet.GetAreasOfXmlMapQuery(queryPath, xmlMap);
            if (linkedAreas.Length > 0)
            {
                Console.WriteLine($"Found {linkedAreas.Length} cell area(s) linked to path '{queryPath}':");
                foreach (CellArea area in linkedAreas)
                {
                    // The area may span multiple rows; iterate through each cell.
                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        for (int col = area.StartColumn; col <= area.EndColumn; col++)
                        {
                            Cell cell = worksheet.Cells[row, col];
                            Console.WriteLine($"  Cell {cell.Name}: \"{cell.StringValue}\"");
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine($"No cells linked to path '{queryPath}'.");
            }

            // ------------------------------------------------------------
            // 4. Save the workbook to verify the result.
            // ------------------------------------------------------------
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
