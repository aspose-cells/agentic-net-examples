// Title: Processing a large XLSX workbook with Aspose.Cells LightCells API and LoadOptions to minimize memory consumption in C#
// Common Searches: Aspose.Cells LightCells example for streaming a 500 MB Excel file in C# | C# load large XLSX with LoadOptions MemoryPreference to keep memory usage low | How to iterate rows and modify values using LightCells API in .NET | Transform first two columns while reading big workbook with Aspose.Cells | Save transformed data to a new XLSX file after LightCells processing in C#
// Tags: LightCells API row‑by‑row processing | LoadOptions memory‑setting configuration Aspose.Cells | memory‑optimized large XLSX handling | Excel numeric value transformation Aspose.Cells | saving processed workbook as Xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a large XLSX file using Aspose.Cells LoadOptions with MemorySetting.MemoryPreference, iterates each used row with the LightCells API, doubles numeric values in the first two columns, writes them to a new workbook, and saves the result as an XLSX file while keeping memory usage low.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the large source workbook
            string sourcePath = "large_input.xlsx";

            // Path where the processed workbook will be saved
            string destPath = "processed_output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook with memory‑optimized settings
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                MemorySetting = MemorySetting.MemoryPreference
            };
            Workbook sourceWorkbook = new Workbook(sourcePath, loadOptions);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new workbook for the processed data
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Determine the range of used rows
            int maxRow = sourceSheet.Cells.MaxDataRow;
            int maxCol = Math.Max(sourceSheet.Cells.MaxDataColumn, 1); // ensure at least two columns

            // Iterate through each row of the source worksheet
            for (int row = 0; row <= maxRow; row++)
            {
                try
                {
                    // Read first two columns (index 0 and 1)
                    object value0 = sourceSheet.Cells[row, 0].Value;
                    object value1 = sourceSheet.Cells[row, 1].Value;

                    // Example transformation: double numeric values
                    if (value0 is double d0) value0 = d0 * 2;
                    if (value1 is double d1) value1 = d1 * 2;

                    // Write transformed values to destination worksheet
                    destSheet.Cells[row, 0].PutValue(value0);
                    destSheet.Cells[row, 1].PutValue(value1);
                }
                catch (Exception rowEx)
                {
                    // Log row‑level errors and continue processing
                    Console.WriteLine($"Row {row} processing error: {rowEx.Message}");
                }
            }

            // Ensure the destination directory exists
            string destDir = Path.GetDirectoryName(destPath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save the processed workbook to disk
            destWorkbook.Save(destPath, SaveFormat.Xlsx);
            Console.WriteLine($"Processed workbook saved to: {destPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
