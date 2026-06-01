using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeValidate
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook that will hold the merged content
                Workbook mergedWorkbook = new Workbook();

                // Paths of workbooks to be merged
                string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

                // Merge each source workbook into the target workbook
                foreach (string file in sourceFiles)
                {
                    if (!File.Exists(file))
                    {
                        Console.WriteLine($"Warning: Source file \"{file}\" not found. Skipping.");
                        continue;
                    }

                    // Load source workbook
                    Workbook src = new Workbook(file);

                    // Copy each worksheet from source to the merged workbook
                    foreach (Worksheet srcSheet in src.Worksheets)
                    {
                        // Add a copy of the source worksheet to the merged workbook
                        mergedWorkbook.Worksheets.AddCopy(srcSheet.Name);
                    }
                }

                // Calculate all formulas in the merged workbook
                mergedWorkbook.CalculateFormula();

                // Validate that no cell contains a #REF! error after merging
                bool hasRefError = false;
                foreach (Worksheet sheet in mergedWorkbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell != null && cell.StringValue == "#REF!")
                            {
                                hasRefError = true;
                                Console.WriteLine($"#REF! error found in sheet '{sheet.Name}' at cell {cell.Name}");
                            }
                        }
                    }
                }

                if (!hasRefError)
                {
                    Console.WriteLine("Validation passed: No #REF! errors found after merging.");
                }

                // Save the merged workbook
                mergedWorkbook.Save("MergedOutput.xlsx");
                Console.WriteLine("Merged workbook saved as \"MergedOutput.xlsx\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}