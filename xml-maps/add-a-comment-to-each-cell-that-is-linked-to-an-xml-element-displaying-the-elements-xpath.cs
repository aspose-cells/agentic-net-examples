// Title: Add XPath comments to Excel cells linked via an XML map using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add comment showing XML element XPath for mapped cells | How to retrieve cell areas of an XML map query and annotate them in Excel with Aspose | Aspose.Cells .NET add visible comment with XPath to cells bound to an XML schema | Iterate XML maps in a workbook and insert XPath comments using Aspose.Cells C# | Save Excel workbook with XPath comments for XML mapped cells using Aspose.Cells
// Tags: annotate xml mapped cells with xpath aspose.cells | getareasofxmlmapquery example aspose.cells c# | xml map cell binding comment aspose.cells | visible excel comment containing xpath aspose.cells | import xml data bind cells aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace XmlMapCommentDemo
{
    // The example loads or creates an Excel workbook with an XML map, defines known XPaths, iterates through each worksheet and XML map, uses GetAreasOfXmlMapQuery to locate cells bound to each XPath, adds a visible comment containing the XPath to those cells, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths for input and output workbooks.
                string inputPath = "InputWithXmlMap.xlsx";
                string outputPath = "OutputWithComments.xlsx";

                Workbook workbook;

                // Load existing workbook if present; otherwise create a sample one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = CreateSampleWorkbookWithXmlMap();
                }

                // For the demo schema we know the relevant XPaths.
                // In a real scenario you would extract them from the schema.
                string[] knownXPaths = new[]
                {
                    "/Root/Item/Name",
                    "/Root/Item/Value"
                };

                // Iterate through worksheets and add comments for linked XML cells.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    for (int mapIdx = 0; mapIdx < workbook.Worksheets.XmlMaps.Count; mapIdx++)
                    {
                        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIdx];

                        foreach (string xpath in knownXPaths)
                        {
                            // Get all cell areas linked to the current XPath.
                            CellArea[] areas = sheet.GetAreasOfXmlMapQuery(xpath, xmlMap);
                            if (areas == null || areas.Length == 0)
                                continue;

                            foreach (CellArea area in areas)
                            {
                                for (int row = area.StartRow; row <= area.EndRow; row++)
                                {
                                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                                    {
                                        string cellName = CellsHelper.CellIndexToName(row, col);
                                        int commentIndex = sheet.Comments.Add(cellName);
                                        Comment comment = sheet.Comments[commentIndex];
                                        comment.Note = $"XPath: {xpath}";
                                        comment.IsVisible = true;
                                    }
                                }
                            }
                        }
                    }
                }

                // Save the workbook with added comments.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with comments at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static Workbook CreateSampleWorkbookWithXmlMap()
        {
            // Sample XML schema (XSD) defining a simple structure.
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                  <xs:element name='Root'>
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

            // Sample XML data that conforms to the above schema.
            string xmlData = @"
                <Root>
                  <Item>
                    <Name>Apple</Name>
                    <Value>10</Value>
                  </Item>
                  <Item>
                    <Name>Banana</Name>
                    <Value>20</Value>
                  </Item>
                </Root>";

            // Create a new workbook and add the XML map.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            int mapIndex = wb.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];

            // Import XML data using the map (creates the cell bindings).
            wb.ImportXml(xmlData, ws.Name, 0, 0);

            return wb;
        }

        // Helper to extract attribute values from a tag (kept for completeness).
        private static string GetAttributeValue(string tag, string attributeName)
        {
            string[] patterns = { $"{attributeName}='", $"{attributeName}=\"" };
            foreach (var pattern in patterns)
            {
                int start = tag.IndexOf(pattern, StringComparison.Ordinal);
                if (start >= 0)
                {
                    start += pattern.Length;
                    int end = tag.IndexOfAny(new[] { '\'', '\"' }, start);
                    if (end > start)
                    {
                        return tag.Substring(start, end - start);
                    }
                }
            }
            return null;
        }
    }
}
