using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetRenameDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and rename them based on data dimensions
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of last used row
                    int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of last used column

                    int sum = maxRow + maxCol;
                    string baseName = $"Sum_{sum}";
                    string safeName = CellsHelper.CreateSafeSheetName(baseName);

                    // Resolve duplicate names (case‑insensitive)
                    string finalName = safeName;
                    int duplicateIndex = 1;
                    while (IsWorksheetNameExists(workbook.Worksheets, finalName, sheet.Index))
                    {
                        finalName = $"{safeName}_{duplicateIndex}";
                        finalName = CellsHelper.CreateSafeSheetName(finalName);
                        duplicateIndex++;
                    }

                    sheet.Name = finalName;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Checks if a worksheet name already exists in the collection, excluding the current sheet.
        private static bool IsWorksheetNameExists(WorksheetCollection sheets, string name, int excludeIndex)
        {
            foreach (Worksheet ws in sheets)
            {
                if (ws.Index != excludeIndex && string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }
}