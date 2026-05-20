using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace XmlMapConflictResolver
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWithXmlMaps.xlsx";
                const string outputPath = "OutputResolved.xlsx";

                // Verify that the input workbook exists before attempting to load it.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load an existing workbook that contains two XML maps.
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook contains at least two XML maps.
                if (workbook.Worksheets.XmlMaps.Count < 2)
                {
                    Console.WriteLine("The workbook must contain at least two XML maps.");
                    return;
                }

                // Access the first worksheet (assumed to contain the mapped cells).
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the two XML maps from the collection.
                XmlMap map1 = workbook.Worksheets.XmlMaps[0];
                XmlMap map2 = workbook.Worksheets.XmlMaps[1];

                // Define the XML element paths that need to be checked for overlapping assignments.
                string[] pathsToCheck = new string[]
                {
                    "/Root/Item",
                    "/Root/Details/Detail",
                    "/Root/Info"
                };

                // Iterate through each path and compare the cell areas linked to both maps.
                foreach (string path in pathsToCheck)
                {
                    // Query cell areas linked to the current path for each map.
                    ArrayList areasMap1 = sheet.XmlMapQuery(path, map1);
                    ArrayList areasMap2 = sheet.XmlMapQuery(path, map2);

                    // Compare every area from map1 with every area from map2.
                    foreach (CellArea area1 in areasMap1)
                    {
                        foreach (CellArea area2 in areasMap2)
                        {
                            if (AreasOverlap(area1, area2))
                            {
                                // Conflict detected – resolve it.
                                // Example strategy: keep the mapping from map1 and clear the overlapping cells from map2.
                                ClearCellArea(sheet, area2);
                                Console.WriteLine(
                                    $"Conflict on path '{path}' resolved: cleared cells of map2 at rows {area2.StartRow}-{area2.EndRow}, cols {area2.StartColumn}-{area2.EndColumn}.");
                            }
                        }
                    }
                }

                // Save the workbook after conflict resolution.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Determines whether two CellArea objects intersect.
        private static bool AreasOverlap(CellArea a, CellArea b)
        {
            bool rowsOverlap = a.StartRow <= b.EndRow && a.EndRow >= b.StartRow;
            bool colsOverlap = a.StartColumn <= b.EndColumn && a.EndColumn >= b.StartColumn;
            return rowsOverlap && colsOverlap;
        }

        // Clears the contents of a given CellArea in the worksheet.
        private static void ClearCellArea(Worksheet sheet, CellArea area)
        {
            // The ClearContents method removes values but keeps formatting.
            sheet.Cells.ClearContents(area.StartRow, area.StartColumn, area.EndRow, area.EndColumn);
        }
    }
}