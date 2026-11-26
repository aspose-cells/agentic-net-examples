// Title: Retrieve the A1 address of the first cell mapped to the /Invoice/Total element using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get cell address for XML map element '/Invoice/Total' | How to find the first worksheet cell mapped to an XML element with Aspose.Cells | C# GetAreasOfXmlMapQuery example for locating Invoice Total cell | Retrieve mapped cell address after importing XML into a workbook using Aspose.Cells | Aspose.Cells GetAreasOfXmlMapQuery returns empty for '/Invoice/Total' XPath
// Tags: GetAreasOfXmlMapQuery usage Aspose.Cells | XML map element to worksheet cell address .NET | import XML into worksheet Aspose.Cells | first mapped cell extraction Aspose.Cells | convert CellArea to A1 style address Aspose.Cells

using System;
using Aspose.Cells;

namespace InvoiceTotalCellAddress
{
    // The example creates a Workbook, imports an XML invoice (which automatically creates an XmlMap), obtains the XmlMap, calls GetAreasOfXmlMapQuery with the XPath '/Invoice/Total' to get the mapped CellArea array, extracts the start row and column of the first area, converts it to an A1‑style address, prints the address, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Sample XML containing an Invoice with a Total element
            string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<Invoice>
    <Header>
        <Date>2023-01-01</Date>
    </Header>
    <Items>
        <Item>
            <Description>Item1</Description>
            <Price>10</Price>
        </Item>
        <Item>
            <Description>Item2</Description>
            <Price>20</Price>
        </Item>
    </Items>
    <Total>30</Total>
</Invoice>";

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Import the XML data into the first worksheet.
                // This automatically creates an XmlMap for the imported XML.
                workbook.ImportXml(xml, "Sheet1", 0, 0);

                // Access the first worksheet (the one we just imported into)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure that an XmlMap was created
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps were created from the imported XML.");
                    return;
                }

                // Retrieve the first (and only) XmlMap
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Query the worksheet for cell areas mapped to the /Invoice/Total element
                // GetAreasOfXmlMapQuery returns an array of CellArea structures.
                CellArea[] mappedAreas = worksheet.GetAreasOfXmlMapQuery("/Invoice/Total", xmlMap);

                if (mappedAreas == null || mappedAreas.Length == 0)
                {
                    Console.WriteLine("No cells are mapped to the '/Invoice/Total' element.");
                }
                else
                {
                    // Take the first mapped area (usually a single cell)
                    CellArea firstArea = mappedAreas[0];

                    // Convert the row/column indices to an A1 style address
                    string cellAddress = worksheet.Cells[firstArea.StartRow, firstArea.StartColumn].Name;

                    Console.WriteLine($"The first cell mapped to '/Invoice/Total' is: {cellAddress}");
                }

                // Optional: Save the workbook to verify the mapping visually
                workbook.Save("InvoiceMappingDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
