// Title: Highlight cells mapped to an XML element path with Worksheet.GetAreasOfXmlMapQuery in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells apply background color to cells returned by GetAreasOfXmlMapQuery | format cells based on XML map query result in a .NET workbook | set solid fill for cells linked to an XML map path using Aspose.Cells | how to change style of XML mapped cells in C# with Aspose.Cells | apply yellow background to cells mapped by an XML element in Aspose.Cells
// Tags: GetAreasOfXmlMapQuery cell formatting | XML map path based cell styling | solid yellow fill Aspose.Cells | C# workbook XML map background color | highlight mapped cells using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace XmlMapBackgroundFormatting
{
    // The example imports an XML string with a namespace into a new workbook, automatically creates an XmlMap, queries the worksheet for all cell areas mapped to a specific XML element path using GetAreasOfXmlMapQuery, and then applies a solid yellow background to each cell in the returned areas before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Sample XML with a namespace. The XML will be imported into the workbook,
                // which automatically creates an XmlMap.
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<ns1:Root xmlns:ns1='http://example.com'>
    <ns1:Data>
        <ns1:Item>Value1</ns1:Item>
        <ns1:Item>Value2</ns1:Item>
        <ns1:Item>Value3</ns1:Item>
    </ns1:Data>
</ns1:Root>";

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Import the XML string using a MemoryStream (ImportXml overload that accepts a Stream).
                using (MemoryStream xmlStream = new MemoryStream())
                using (StreamWriter writer = new StreamWriter(xmlStream))
                {
                    writer.Write(xml);
                    writer.Flush();
                    xmlStream.Position = 0; // Reset stream position before import.

                    workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
                }

                // Ensure that an XmlMap was created.
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XmlMap was created after importing XML.");
                    return;
                }

                // Retrieve the first (and only) XmlMap that was created.
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Reference to the worksheet that now contains the imported XML data.
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the XML element path we want to query.
                // The path must match the namespace prefixes used in the XML.
                string queryPath = "/ns1:Root/ns1:Data/ns1:Item";

                // Query the worksheet for all cell areas that are mapped to the specified path.
                CellArea[] mappedAreas = worksheet.GetAreasOfXmlMapQuery(queryPath, xmlMap);

                if (mappedAreas == null || mappedAreas.Length == 0)
                {
                    Console.WriteLine($"No cells were mapped to the path '{queryPath}'.");
                }
                else
                {
                    Console.WriteLine($"Found {mappedAreas.Length} mapped area(s). Applying background color...");

                    // Apply a solid yellow background to every cell in each mapped area.
                    foreach (CellArea area in mappedAreas)
                    {
                        for (int row = area.StartRow; row <= area.EndRow; row++)
                        {
                            for (int col = area.StartColumn; col <= area.EndColumn; col++)
                            {
                                // Get the current style, modify it, and set it back.
                                Style style = worksheet.Cells[row, col].GetStyle();
                                style.ForegroundColor = Color.Yellow;
                                style.Pattern = BackgroundType.Solid;
                                worksheet.Cells[row, col].SetStyle(style);
                            }
                        }
                    }

                    Console.WriteLine("Background color applied to all mapped cells.");
                }

                // Save the workbook to verify the result.
                string outputPath = "XmlMapBackgroundFormatted.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
