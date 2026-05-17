using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapClear
{
    public class XmlMapCleaner
    {
        public static void ClearLinkedCells(string workbookPath, string outputPath, string xmlMapName, string xmlPath)
        {
            // Ensure the source workbook exists; create an empty one if it does not.
            if (!File.Exists(workbookPath))
            {
                var emptyWb = new Workbook();
                emptyWb.Save(workbookPath);
            }

            Workbook workbook = new Workbook(workbookPath);

            XmlMap targetMap = null;
            foreach (XmlMap map in workbook.Worksheets.XmlMaps)
            {
                if (map.Name == xmlMapName)
                {
                    targetMap = map;
                    break;
                }
            }

            if (targetMap == null)
            {
                Console.WriteLine($"XML map '{xmlMapName}' not found.");
                return;
            }

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ArrayList linkedAreas = sheet.XmlMapQuery(xmlPath, targetMap);
                if (linkedAreas != null)
                {
                    foreach (CellArea area in linkedAreas)
                    {
                        sheet.Cells.ClearContents(area);
                    }
                }
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Linked cells cleared and workbook saved to '{outputPath}'.");
        }

        public static void Run()
        {
            string inputFile = "InputWithXmlMap.xlsx";
            string outputFile = "ClearedXmlMap.xlsx";
            string mapName = "Transmittals_Map";
            string path = "/Transmittals/Issued_Document";

            ClearLinkedCells(inputFile, outputFile, mapName, path);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            XmlMapCleaner.Run();
        }
    }
}