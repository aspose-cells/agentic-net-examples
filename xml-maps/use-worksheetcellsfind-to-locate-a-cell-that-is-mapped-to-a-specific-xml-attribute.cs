// Title: Find a cell mapped to a specific XML attribute in an Excel worksheet using Aspose.Cells C# (Worksheet.Cells.Find)
// Common Searches: c# find cell that corresponds to an XML attribute using Aspose.Cells | how to use GetAreasOfXmlMapQuery with Worksheet.Cells.Find in Aspose.Cells | search for XML attribute id value in Excel sheet with Aspose.Cells .NET | locate mapped XML attribute cell in workbook programmatically Aspose.Cells | retrieve cell address for XML attribute value using Aspose.Cells C#
// Tags: Worksheet.Cells.Find with XML map | GetAreasOfXmlMapQuery attribute mapping | search XML attribute values in Excel workbook | map XML attribute to worksheet cells | find cell by XML attribute id Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlAttributeFind
{
    // The example creates sample XML and XSD files, adds an XML map to a new workbook, imports the XML data, obtains the cell areas linked to the '/Root/Item/@id' attribute via GetAreasOfXmlMapQuery, and then uses Worksheet.Cells.Find with appropriate FindOptions to locate the cell containing the attribute value "102". It confirms the cell belongs to a mapped area and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Prepare sample XML and XSD files (written to the execution folder)
                string xmlContent = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Item id='101'>Apple</Item>
    <Item id='102'>Orange</Item>
    <Item id='103'>Banana</Item>
</Root>";

                string xsdContent = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:simpleContent>
              <xs:extension base='xs:string'>
                <xs:attribute name='id' type='xs:string' use='required' />
              </xs:extension>
            </xs:simpleContent>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                string xmlPath = Path.Combine(Environment.CurrentDirectory, "sample.xml");
                string xsdPath = Path.Combine(Environment.CurrentDirectory, "sample.xsd");

                // Write files (overwrites if they already exist)
                File.WriteAllText(xmlPath, xmlContent);
                File.WriteAllText(xsdPath, xsdContent);

                // Verify that the required files exist
                if (!File.Exists(xmlPath))
                {
                    Console.WriteLine($"XML file not found: {xmlPath}");
                    return;
                }

                if (!File.Exists(xsdPath))
                {
                    Console.WriteLine($"XSD file not found: {xsdPath}");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add the XML map based on the XSD schema
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xsdPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Import the XML data into the first worksheet (Sheet1)
                workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the XML attribute path we are interested in
                string attributePath = "/Root/Item/@id";

                // Query the worksheet for all cell areas that are mapped to the specified attribute
                CellArea[] mappedAreas = worksheet.GetAreasOfXmlMapQuery(attributePath, xmlMap);

                Console.WriteLine($"Number of cell areas mapped to attribute '{attributePath}': {mappedAreas.Length}");
                foreach (var area in mappedAreas)
                {
                    Console.WriteLine($"Mapped Area: Rows {area.StartRow}-{area.EndRow}, Columns {area.StartColumn}-{area.EndColumn}");
                }

                // The value we want to locate – for example, the attribute value "102"
                string valueToFind = "102";

                // Configure find options (search whole sheet, whole cell content)
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values
                    // The default LookAtType performs a whole‑cell match; no explicit enum value is required.
                };

                // Perform the search
                Cell foundCell = worksheet.Cells.Find(valueToFind, null, findOptions);

                if (foundCell != null)
                {
                    Console.WriteLine($"Found value '{valueToFind}' in cell {foundCell.Name} (Row {foundCell.Row}, Column {foundCell.Column})");

                    // Verify that the found cell belongs to one of the mapped areas
                    bool isMapped = IsCellInAnyArea(foundCell.Row, foundCell.Column, mappedAreas);
                    Console.WriteLine(isMapped
                        ? "The found cell is correctly mapped to the specified XML attribute."
                        : "The found cell is NOT mapped to the specified XML attribute.");
                }
                else
                {
                    Console.WriteLine($"Value '{valueToFind}' was not found in the worksheet.");
                }

                // Save the workbook for visual inspection (optional)
                string outputPath = Path.Combine(Environment.CurrentDirectory, "MappedWorkbook.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static bool IsCellInAnyArea(int row, int column, CellArea[] areas)
        {
            foreach (var area in areas)
            {
                if (row >= area.StartRow && row <= area.EndRow &&
                    column >= area.StartColumn && column <= area.EndColumn)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
