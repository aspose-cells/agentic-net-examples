// Title: Retrieve cell addresses mapped to a namespaced XPath using Worksheet.XmlMapQuery in Aspose.Cells for C#
// Common Searches: Aspose.Cells C# get cells linked to an XPath after ImportXml | Worksheet.XmlMapQuery example with namespace prefixes in .NET | How to list cell addresses for a specific XML map XPath using Aspose.Cells | C# code to retrieve CellArea objects for an XPath in an Excel workbook | Query mapped cells for /ns1:Root/ns1:Data/ns1:Item using Aspose.Cells
// Tags: Worksheet.XmlMapQuery usage | XPath cell mapping Aspose.Cells | ImportXml with namespace Aspose.Cells | CellArea enumeration from XML map | C# Aspose.Cells XML map query

using System;
using System.Collections;
using Aspose.Cells;

namespace XmlMapQueryDemo
{
    // The example imports a namespaced XML document into a new workbook, automatically creates an XML map, and then uses Worksheet.XmlMapQuery with a defined XPath to obtain the CellArea objects that correspond to the mapped cells. It iterates through each cell in the returned areas, prints their addresses and values, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Sample XML with a namespace
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<ns1:Root xmlns:ns1='http://example.com'>
    <ns1:Data>
        <ns1:Item>Value1</ns1:Item>
        <ns1:Item>Value2</ns1:Item>
    </ns1:Data>
</ns1:Root>";

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Import the XML into the first worksheet starting at cell A1.
                // This also creates an XML map and links the XML elements to cells.
                workbook.ImportXml(xml, "Sheet1", 0, 0);

                // Access the first worksheet (the one we just imported into)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the XML map that was created during ImportXml
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps were created. Exiting.");
                    return;
                }
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Define the XPath expression we want to query.
                // Note: The namespace prefix must match the one defined in the XML.
                string xpath = "/ns1:Root/ns1:Data/ns1:Item";

                // Query the worksheet for cell areas linked to the specified XPath.
                // XmlMapQuery returns an ArrayList of CellArea objects.
                ArrayList cellAreas = worksheet.XmlMapQuery(xpath, xmlMap);

                // Output the results
                if (cellAreas.Count == 0)
                {
                    Console.WriteLine($"No cells are mapped to the XPath '{xpath}'.");
                }
                else
                {
                    Console.WriteLine($"Cells mapped to XPath '{xpath}':");
                    foreach (object obj in cellAreas)
                    {
                        if (obj is CellArea area)
                        {
                            // The mapping may cover a range; iterate through each cell in the area.
                            for (int row = area.StartRow; row <= area.EndRow; row++)
                            {
                                for (int col = area.StartColumn; col <= area.EndColumn; col++)
                                {
                                    Cell cell = worksheet.Cells[row, col];
                                    string address = CellsHelper.CellIndexToName(row, col);
                                    Console.WriteLine($"  {address} => \"{cell.StringValue}\" (Row {row}, Column {col})");
                                }
                            }
                        }
                    }
                }

                // Save the workbook so you can inspect the mapping in Excel if desired.
                string outputPath = "XmlMapQueryDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
