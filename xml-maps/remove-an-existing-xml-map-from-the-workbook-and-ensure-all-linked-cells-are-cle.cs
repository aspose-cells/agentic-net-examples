using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsXmlMapRemoval
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one XML map to remove
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                // Get the first XML map (you can select a different one by index or name)
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Build a root path for querying linked cells (e.g., "/RootElement")
                string rootPath = "/" + xmlMap.RootElementName;

                // Iterate through all worksheets and clear cells linked to the XML map
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Query cell areas that are linked to the specified path of the XML map
                    ArrayList linkedAreas = sheet.XmlMapQuery(rootPath, xmlMap);

                    // Clear each linked cell area
                    foreach (CellArea area in linkedAreas)
                    {
                        for (int row = area.StartRow; row <= area.EndRow; row++)
                        {
                            for (int col = area.StartColumn; col <= area.EndColumn; col++)
                            {
                                sheet.Cells[row, col].PutValue(string.Empty);
                            }
                        }
                    }
                }

                // Remove the XML map from the collection
                workbook.Worksheets.XmlMaps.RemoveAt(0);
            }

            // Save the modified workbook (replace with your desired output path)
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}