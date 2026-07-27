using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapRemoval
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Remove all XML maps from the workbook
                workbook.Worksheets.XmlMaps.Clear();

                // Clear the contents of all cells in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxColumn = sheet.Cells.MaxDataColumn;

                    // Skip empty sheets
                    if (maxRow < 0 || maxColumn < 0)
                        continue;

                    // Define the area covering all used cells
                    CellArea area = new CellArea
                    {
                        StartRow = 0,
                        StartColumn = 0,
                        EndRow = maxRow,
                        EndColumn = maxColumn
                    };

                    sheet.Cells.ClearContents(area);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}