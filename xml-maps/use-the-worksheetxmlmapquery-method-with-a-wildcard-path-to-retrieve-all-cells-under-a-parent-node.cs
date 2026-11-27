// Title: Using Worksheet.XmlMapQuery with a wildcard path to get all child cells of an XML element in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# query XML map using wildcard to list child node cells | How to retrieve cells mapped to any child of an XML parent element with Worksheet.XmlMapQuery | C# example of XmlMapQuery wildcard path "/ns:Root/ns:Items/*" in Aspose.Cells | Get CellArea collection for XML child nodes in a workbook using Aspose.Cells | Import XML into Aspose.Cells and query all mapped cells under a parent node
// Tags: Worksheet XmlMapQuery wildcard path usage | Aspose.Cells XML map child node extraction | C# retrieve CellArea list from XmlMapQuery | Import XML into workbook with Aspose.Cells | Processing XmlMap query results in .NET | Saving workbook after XML map query Aspose.Cells

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// The example loads XML from a file or an embedded string, imports it into a new workbook (which automatically creates an XmlMap), obtains the first XmlMap, and calls Worksheet.XmlMapQuery with the wildcard path "/ns:Root/ns:Items/*" to collect all CellArea objects that correspond to any child of the Items node. It then iterates the results, prints each cell address and value, and saves the workbook as XmlMapQueryResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Prepare XML source (file or embedded string)
            // ------------------------------------------------------------
            string xmlPath = "sample.xml";
            string xmlContent;

            if (File.Exists(xmlPath))
            {
                xmlContent = File.ReadAllText(xmlPath);
                Console.WriteLine($"Loaded XML from file: {xmlPath}");
            }
            else
            {
                // Fallback XML string with a parent node <Items> containing several <Item> elements
                xmlContent = @"<?xml version='1.0' encoding='UTF-8'?>
<ns:Root xmlns:ns='http://example.com'>
    <ns:Items>
        <ns:Item>Alpha</ns:Item>
        <ns:Item>Beta</ns:Item>
        <ns:Item>Gamma</ns:Item>
    </ns:Items>
</ns:Root>";
                Console.WriteLine("XML file not found. Using embedded XML string.");
            }

            // ------------------------------------------------------------
            // 2. Create a new workbook and import the XML.
            //    ImportXml creates an XmlMap automatically.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Import the XML starting at cell A1 (row 0, column 0)
            workbook.ImportXml(xmlContent, "Sheet1", 0, 0);
            Console.WriteLine("XML imported into worksheet.");

            // ------------------------------------------------------------
            // 3. Retrieve the XmlMap that was created during import.
            // ------------------------------------------------------------
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps were created. Exiting.");
                return;
            }

            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
            Console.WriteLine($"XmlMap name: {xmlMap.Name}");
            Console.WriteLine($"Root element: {xmlMap.RootElementName}");

            // ------------------------------------------------------------
            // 4. Query all cells that are mapped to any child of the parent node.
            //    Wildcard path: "/ns:Root/ns:Items/*"
            // ------------------------------------------------------------
            string wildcardPath = "/ns:Root/ns:Items/*";
            Console.WriteLine($"Executing XmlMapQuery with wildcard path: {wildcardPath}");

            // NOTE: XmlMapQuery is marked obsolete but still functional.
            ArrayList cellAreas = worksheet.XmlMapQuery(wildcardPath, xmlMap);

            // ------------------------------------------------------------
            // 5. Process and display the results.
            // ------------------------------------------------------------
            if (cellAreas.Count == 0)
            {
                Console.WriteLine("No cells were found for the specified wildcard path.");
            }
            else
            {
                Console.WriteLine($"Found {cellAreas.Count} cell area(s):");
                foreach (object obj in cellAreas)
                {
                    if (obj is CellArea area)
                    {
                        // Retrieve the cell value (assuming a single cell per area)
                        string cellName = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                        string cellValue = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;

                        Console.WriteLine($"- Cell {cellName} (Row {area.StartRow}, Column {area.StartColumn}) = \"{cellValue}\"");
                    }
                }
            }

            // ------------------------------------------------------------
            // 6. Save the workbook for visual verification (optional).
            // ------------------------------------------------------------
            string outputFile = "XmlMapQueryResult.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
