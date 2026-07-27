using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapClear
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                XmlMapHelper.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error: {ex.Message}");
            }
        }
    }

    public class XmlMapHelper
    {
        /// <summary>
        /// Clears the contents of all cells linked to the specified XML map and path.
        /// The XML map itself remains intact.
        /// </summary>
        public static void ClearLinkedCells(Workbook workbook, string mapName, string xmlPath)
        {
            // Locate the XmlMap by name
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
                Console.WriteLine($"XmlMap with name '{mapName}' not found.");
                return;
            }

            // Iterate through each worksheet and clear linked cell areas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // XmlMapQuery returns an ArrayList of CellArea objects
                ArrayList linkedAreas = sheet.XmlMapQuery(xmlPath, targetMap);
                foreach (CellArea area in linkedAreas)
                {
                    // Clear only the contents; formatting and the link remain
                    sheet.Cells.ClearContents(area);
                }
            }
        }

        // Example usage
        public static void Run()
        {
            string inputPath = "InputWithXmlMap.xlsx";
            string outputPath = "OutputCleared.xlsx";

            // Prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            Workbook wb;
            try
            {
                wb = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            string mapName = "Transmittals_Map";
            string xmlPath = "/Transmittals/Issued_Document";

            try
            {
                ClearLinkedCells(wb, mapName, xmlPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error clearing linked cells: {ex.Message}");
            }

            try
            {
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}