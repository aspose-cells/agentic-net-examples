// Title: C# – Filter XML‑mapped cells by XPath using Aspose.Cells
// Description: Demonstrates how to import namespaced XML into a workbook, obtain the XML map collection across Aspose.Cells versions, query for CellArea ranges that match a namespace‑aware XPath, iterate the cells to read values, and save the result.
// Keywords: Aspose.Cells | C# | .NET | XML map | XPath query | CellArea | namespace aware XML | version‑agnostic API | filter mapped cells | import XML workbook
// Common Searches: Aspose.Cells query cells by XPath | Get CellArea for XML map element C# | How to filter XML‑mapped ranges in Aspose.Cells | Version‑independent XML map collection Aspose.Cells | Read values of XML‑mapped cells using XPath
// Developer Intent: Select and process only the worksheet cells linked to a specific XPath in an XML map.
// Use Cases: Extract values of all <Item> nodes from an imported XML map by iterating over the corresponding CellArea ranges. | Create a report that operates solely on cells mapped to a particular XML element while ignoring unrelated data. | Update or replace values in cells mapped to a given XPath and save the modified workbook.
// AI Prompts: Generate C# code with Aspose.Cells that imports a namespaced XML file, queries the worksheet for CellArea objects matching a specific XPath, and prints each cell's value. | Show a version‑agnostic method to obtain the XML map collection in Aspose.Cells and filter mapped areas using a namespace‑aware XPath expression. | Explain how to modify the sample to write new data back to the cells that match the XPath and then save the workbook.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Demonstrates how to import namespaced XML into a workbook, obtain the XML map collection across Aspose.Cells versions, query for CellArea ranges that match a namespace‑aware XPath, iterate the cells to read values, and save the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample XML with a namespace; it will be imported as an XML map
            string xml = @"<?xml version='1.0' encoding='UTF-8'?>
                <ns1:Root xmlns:ns1='http://example.com'>
                    <ns1:Data>
                        <ns1:Item>Value1</ns1:Item>
                        <ns1:Item>Value2</ns1:Item>
                    </ns1:Data>
                </ns1:Root>";

            // Import the XML into the worksheet starting at cell A1; this creates an XML map
            workbook.ImportXml(xml, "Sheet1", 0, 0);

            // Access the collection of XML maps (property name may vary by version)
            // Use XmlMaps if available; otherwise fall back to XmlMapCollection
            var xmlMapCollection = GetXmlMapCollection(workbook);
            if (xmlMapCollection == null || xmlMapCollection.Count == 0)
            {
                Console.WriteLine("No XML map was created.");
                return;
            }

            // Retrieve the first XML map
            XmlMap xmlMap = xmlMapCollection[0] as XmlMap;

            // Define the XPath expression to locate the desired XML elements
            string xpath = "/ns1:Root/ns1:Data/ns1:Item";

            // Query the worksheet for cell areas that are mapped to the specified XPath
            ArrayList mappedAreas = worksheet.XmlMapQuery(xpath, xmlMap);

            // Process each returned CellArea
            foreach (CellArea area in mappedAreas)
            {
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        Console.WriteLine($"Cell {cell.Name}: {cell.StringValue}");
                    }
                }
            }

            // Save the workbook (ensure the directory is writable)
            string outputPath = "FilteredXmlMap.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper to obtain the XML map collection compatible with different Aspose.Cells versions
    private static XmlMapCollection GetXmlMapCollection(Workbook workbook)
    {
        // Prefer the XmlMaps property if it exists
        var type = typeof(Workbook);
        var prop = type.GetProperty("XmlMaps");
        if (prop != null)
        {
            return prop.GetValue(workbook) as XmlMapCollection;
        }

        // Fallback to XmlMapCollection property (older versions)
        prop = type.GetProperty("XmlMapCollection");
        if (prop != null)
        {
            return prop.GetValue(workbook) as XmlMapCollection;
        }

        return null;
    }
}
