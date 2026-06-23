using System;
using System.IO;
using Aspose.Cells;

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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (LightCells API removed for compatibility)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Retrieve all merged cell areas
            CellArea[] mergedAreas = cells.GetMergedAreas();

            Console.WriteLine($"Number of merged areas: {mergedAreas.Length}");
            foreach (CellArea area in mergedAreas)
            {
                // Convert zero‑based indices to Excel style addresses (e.g., A1)
                string startAddress = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                string endAddress   = CellsHelper.CellIndexToName(area.EndRow,   area.EndColumn);

                // Output start and end coordinates
                Console.WriteLine(
                    $"Merged area: {startAddress}:{endAddress} " +
                    $"(StartRow={area.StartRow}, StartColumn={area.StartColumn}, " +
                    $"EndRow={area.EndRow}, EndColumn={area.EndColumn})");
            }

            // Optional: save the workbook (no changes made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}