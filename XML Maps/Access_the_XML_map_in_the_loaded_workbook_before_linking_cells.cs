using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (uses workbook-load rule)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Access the first worksheet (uses worksheet-access rule)
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the collection of XML maps in the workbook
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Ensure that at least one XML map exists
            if (xmlMaps.Count > 0)
            {
                // Retrieve the first XML map
                XmlMap xmlMap = xmlMaps[0];

                // Example: link cell A1 (row 0, column 0) to a specific XML path
                // This part uses free‑form code because no specific rule exists for LinkToXmlMap
                worksheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item");
            }
            else
            {
                Console.WriteLine("No XML maps found in the workbook.");
            }

            // Save the modified workbook (uses workbook-save rule)
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}