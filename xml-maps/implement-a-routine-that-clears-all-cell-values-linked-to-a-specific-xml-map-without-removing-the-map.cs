// Title: Aspose.Cells C# – Clear Values of Cells Linked to an XML Map Without Deleting the Map
// Description: A C# routine that finds an XmlMap by name, uses XmlMapQuery to locate all cell areas linked to a specified XML path, and clears their contents with Cells.ClearContents while keeping the XML map definition intact in an Aspose.Cells workbook.
// Keywords: Aspose.Cells | C# | XML map | clear linked cells | XmlMapQuery | Cells.ClearContents | preserve XML map | remove mapped data | Excel automation | Aspose.Cells .NET
// Common Searches: clear cells linked to xml map aspose.cells c# | remove xml mapped data without deleting map | aspnet clear xml map linked cells | aspose.cells XmlMapQuery example | how to reset xml mapped workbook
// Developer Intent: Remove all data from cells bound to a specific XML map path while retaining the map definition for future imports.
// Use Cases: Reset a template workbook before loading new XML data, preserving the existing map. | Clean user‑filled forms bound to an XML schema so the file can be reused without recreating the map. | Archive a report workbook by stripping previous XML‑linked values while keeping the mapping structure for later analysis.
// AI Prompts: Write C# code with Aspose.Cells to clear cells linked to an XML map named 'OrdersMap' for the path '/Orders/Order' without deleting the map. | Explain how XmlMapQuery and Cells.ClearContents work together to erase linked data while leaving the XmlMap unchanged. | Add detailed error handling to the ClearLinkedCells method to log missing maps, empty results, and unexpected exceptions.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapClear
{
    // A C# routine that finds an XmlMap by name, uses XmlMapQuery to locate all cell areas linked to a specified XML path, and clears their contents with Cells.ClearContents while keeping the XML map definition intact in an Aspose.Cells workbook.
    public class XmlMapHelper
    {
        /// <param name="workbook">The workbook containing the XML map.</param>
        /// <param name="mapName">The name of the XML map to target.</param>
        /// <param name="xmlPath">The XML element path used when linking cells to the map.</param>
        public static void ClearLinkedCells(Workbook workbook, string mapName, string xmlPath)
        {
            // Locate the XmlMap by its name.
            XmlMap targetMap = null;
            foreach (XmlMap map in workbook.Worksheets.XmlMaps)
            {
                if (map.Name == mapName)
                {
                    targetMap = map;
                    break;
                }
            }

            // If the map is not found, exit the method.
            if (targetMap == null)
                return;

            // Iterate through every worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Query the worksheet for cell areas linked to the given path of the target map.
                ArrayList linkedAreas = sheet.XmlMapQuery(xmlPath, targetMap);

                // Clear the contents of each returned cell area.
                foreach (CellArea area in linkedAreas)
                {
                    sheet.Cells.ClearContents(area);
                }
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            const string inputPath = "InputWithXmlMap.xlsx";
            const string outputPath = "OutputCleared.xlsx";

            try
            {
                Workbook wb;

                // Ensure the input file exists; otherwise create an empty workbook.
                if (File.Exists(inputPath))
                {
                    wb = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                    wb = new Workbook(); // creates a default workbook with one worksheet
                }

                // Specify the XML map name and the path that was used when linking cells.
                string mapName = "MyXmlMap";
                string xmlPath = "/Root/Item";

                // Clear all cell values linked to the specified map/path.
                XmlMapHelper.ClearLinkedCells(wb, mapName, xmlPath);

                // Save the workbook; the XML map itself is preserved if it existed.
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
