using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveXmlMapAndClearLinkedCells
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "InputWithXmlMap.xlsx";
            const string outputPath = "OutputWithoutXmlMap.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Remove all XML maps if any exist
                if (workbook.Worksheets.XmlMaps.Count > 0)
                {
                    for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
                    {
                        XmlMap map = workbook.Worksheets.XmlMaps[i];
                        Console.WriteLine($"Removing XML map: {map.Name}");
                    }

                    workbook.Worksheets.XmlMaps.Clear();
                }

                // Clear contents of every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range; if sheet is empty, skip clearing
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxColumn = sheet.Cells.MaxDataColumn;

                    if (maxRow < 0 || maxColumn < 0)
                        continue; // nothing to clear

                    // Define the area covering all used cells
                    CellArea area = new CellArea
                    {
                        StartRow = 0,
                        StartColumn = 0,
                        EndRow = maxRow,
                        EndColumn = maxColumn
                    };

                    // Clear values, formulas and formatting within the area
                    sheet.Cells.ClearContents(area);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine("XML maps removed and linked cells cleared successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}