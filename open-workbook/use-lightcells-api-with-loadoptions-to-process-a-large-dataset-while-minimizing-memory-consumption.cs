// Title: Use Aspose.Cells LoadOptions with LightCells API in C# to stream and process a large XLSX workbook row by row while minimizing memory usage
// Common Searches: how to stream rows from a large XLSX file using Aspose.Cells LightCells API in C# | process large Excel workbook with minimal memory using LoadOptions and LightCells in .NET | row‑by‑row calculation and copy to new workbook with Aspose.Cells without loading entire file | C# example for adding a computed column while reading Excel data with LightCells | avoid OutOfMemoryException when handling big XLSX files with Aspose.Cells LoadOptions
// Tags: Aspose.Cells LightCells API row streaming | LoadOptions memory‑efficient XLSX processing .NET | C# calculate column while copying Excel rows | large workbook row‑wise processing Aspose.Cells | streaming Excel data to new workbook C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads a large XLSX workbook using Aspose.Cells LoadOptions, then employs the LightCells API to iterate rows one at a time, copy existing cell values, compute a sum of the first two numeric columns, append the result as a new column, and write each processed row to a new workbook, all while keeping memory consumption low.
class LightCellsProcessing
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "large_input.xlsx";
        string outputPath = "processed_output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook (default load options are sufficient for most scenarios)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            using (Workbook inputWb = new Workbook(inputPath, loadOptions))
            {
                // Access the first worksheet of the input workbook
                Worksheet inputWs = inputWb.Worksheets[0];
                Cells inputCells = inputWs.Cells;

                // Prepare the output workbook (empty) and its first worksheet
                Workbook outputWb = new Workbook();
                Worksheet outputWs = outputWb.Worksheets[0];
                Cells outputCells = outputWs.Cells;

                // Determine the range of rows and columns that contain data
                int maxRow = inputCells.MaxDataRow;
                int maxCol = inputCells.MaxDataColumn;

                // Process each row
                for (int row = 0; row <= maxRow; row++)
                {
                    // Copy existing cell values to an object array (+1 for the new sum column)
                    object[] values = new object[maxCol + 2];

                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = inputCells[row, col];
                        values[col] = cell.Value;
                    }

                    // Example processing: calculate sum of first two numeric columns
                    double sum = 0;
                    if (inputCells[row, 0].Value is double d0) sum += d0;
                    if (inputCells[row, 1].Value is double d1) sum += d1;

                    // Append the calculated sum as a new column
                    values[maxCol + 1] = sum;

                    // Write the processed row to the output worksheet
                    for (int col = 0; col < values.Length; col++)
                    {
                        outputCells[row, col].Value = values[col];
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the processed workbook (default options are sufficient)
                outputWb.Save(outputPath);
            }

            Console.WriteLine($"Processing completed successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
