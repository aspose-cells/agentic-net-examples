// Title: C# – Get First Cell Address Mapped to /Invoice/Total Using Aspose.Cells XML Map
// Description: This example shows how to import an XML invoice into an Aspose.Cells workbook, access the generated XmlMap (via reflection), query the worksheet with XmlMapQuery for the "/Invoice/Total" element, and extract the address of the first mapped cell (e.g., "C5"). The code also demonstrates safe handling when no mapping exists and optionally saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | XML map | XmlMapQuery | first mapped cell address | Invoice Total cell | worksheet cell address | reflection XmlMaps | import XML to workbook
// Common Searches: Aspose.Cells get address of cell mapped to XML element | C# XmlMapQuery first cell address | how to retrieve mapped cell for /Invoice/Total | Aspose.Cells XML map example C# | using reflection to access XmlMaps Aspose.Cells
// Developer Intent: Find the address of the first worksheet cell that is linked to the /Invoice/Total element in an XML‑mapped workbook.
// Use Cases: Show the mapped cell location in a UI to verify XML‑to‑cell binding. | Validate that the Total element is correctly linked before performing calculations. | Programmatically update the Total cell after modifying the source XML.
// AI Prompts: Generate C# code that imports XML into an Aspose.Cells workbook and returns the address of the first cell mapped to "/Invoice/Total". | Explain how to handle cases where XmlMapQuery returns an empty result for a given XPath in Aspose.Cells. | Provide a non‑reflection approach to access the XmlMaps collection in Aspose.Cells for .NET.

using System;
using System.Collections;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsXmlMapExample
{
    // This example shows how to import an XML invoice into an Aspose.Cells workbook, access the generated XmlMap (via reflection), query the worksheet with XmlMapQuery for the "/Invoice/Total" element, and extract the address of the first mapped cell (e.g., "C5"). The code also demonstrates safe handling when no mapping exists and optionally saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet.
                Workbook workbook = new Workbook();

                // Reference to the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample XML containing the /Invoice/Total element.
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<Invoice>
    <Header>
        <Date>2023-01-01</Date>
    </Header>
    <Details>
        <Item>Item1</Item>
        <Item>Item2</Item>
    </Details>
    <Total>1234.56</Total>
</Invoice>";

                // Import the XML into the worksheet starting at cell A1.
                // This also creates an XML map inside the workbook.
                workbook.ImportXml(xml, "Sheet1", 0, 0);

                // Use reflection to obtain the XmlMaps collection (avoids compile‑time dependency).
                PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps", BindingFlags.Public | BindingFlags.Instance);
                if (xmlMapsProp == null)
                {
                    Console.WriteLine("The current Aspose.Cells version does not support XML maps.");
                    return;
                }

                dynamic xmlMaps = xmlMapsProp.GetValue(workbook);
                if (xmlMaps == null || xmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps were created.");
                    return;
                }

                // Retrieve the first (and only) XML map.
                dynamic xmlMap = xmlMaps[0];

                // Query the worksheet for cell areas mapped to the /Invoice/Total element.
                ArrayList cellAreas = worksheet.XmlMapQuery("/Invoice/Total", xmlMap);

                if (cellAreas != null && cellAreas.Count > 0)
                {
                    // Get the first mapped area.
                    CellArea area = (CellArea)cellAreas[0];

                    // Obtain the first cell in that area.
                    Cell firstMappedCell = worksheet.Cells[area.StartRow, area.StartColumn];

                    // Retrieve its address (e.g., "C5").
                    string address = firstMappedCell.Name;

                    Console.WriteLine($"First cell mapped to /Invoice/Total: {address}");
                }
                else
                {
                    Console.WriteLine("No cells are mapped to /Invoice/Total.");
                }

                // Save the workbook (optional). Ensure the directory exists.
                string outputPath = "InvoiceMappingDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
