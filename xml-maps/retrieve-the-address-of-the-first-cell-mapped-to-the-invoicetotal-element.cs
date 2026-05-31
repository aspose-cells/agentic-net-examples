using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapQueryDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that (potentially) contains an XML map
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook has at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Error: The workbook does not contain any worksheets.");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Verify that at least one XML map is defined
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("Error: No XML maps are defined in the workbook.");
                    return;
                }

                // Retrieve the first XML map
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Query cells mapped to the specified XML element
                ArrayList cellAreas = worksheet.XmlMapQuery("/Invoice/Total", xmlMap);

                if (cellAreas.Count > 0)
                {
                    // The first CellArea corresponds to the first mapped range
                    CellArea area = (CellArea)cellAreas[0];
                    Cell firstMappedCell = worksheet.Cells[area.StartRow, area.StartColumn];
                    Console.WriteLine("First cell mapped to /Invoice/Total: " + firstMappedCell.Name);
                }
                else
                {
                    Console.WriteLine("No cells are mapped to /Invoice/Total.");
                }

                // Optional: save modifications
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}