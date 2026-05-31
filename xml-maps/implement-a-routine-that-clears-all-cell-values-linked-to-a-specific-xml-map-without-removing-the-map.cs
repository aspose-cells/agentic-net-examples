using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class XmlMapHelper
    {
        /// <summary>
        /// Clears the values of all cells that are linked to the specified XML map.
        /// The XML map itself remains in the workbook.
        /// </summary>
        /// <param name="workbook">The workbook containing the XML map.</param>
        /// <param name="mapName">The name of the XML map whose linked cells should be cleared.</param>
        public static void ClearLinkedCells(Workbook workbook, string mapName)
        {
            try
            {
                // Find the XmlMap by name.
                XmlMap targetMap = null;
                foreach (XmlMap map in workbook.Worksheets.XmlMaps)
                {
                    if (map.Name == mapName)
                    {
                        targetMap = map;
                        break;
                    }
                }

                if (targetMap == null)
                {
                    Console.WriteLine($"XmlMap \"{mapName}\" not found.");
                    return;
                }

                // Build a path that points to the root element of the map.
                string rootPath = "/" + targetMap.RootElementName;

                // Iterate through each worksheet and clear the linked cell areas.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ArrayList linkedAreas = sheet.XmlMapQuery(rootPath, targetMap);
                    foreach (CellArea area in linkedAreas)
                    {
                        sheet.Cells.ClearContents(area);
                    }
                }

                Console.WriteLine($"All cells linked to XmlMap \"{mapName}\" have been cleared.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error clearing linked cells: {ex.Message}");
            }
        }

        // Example usage
        public static void RunExample()
        {
            const string inputPath = "InputWithXmlMap.xlsx";
            const string outputPath = "OutputCleared.xlsx";
            const string mapName = "MyXmlMap";

            // Prevent FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load an existing workbook that contains an XML map.
                Workbook wb = new Workbook(inputPath);

                // Clear linked cells.
                ClearLinkedCells(wb, mapName);

                // Save the workbook (the XML map remains intact).
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            XmlMapHelper.RunExample();
        }
    }
}