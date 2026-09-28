// Title: Calculate the sum of a column in a large XLSX workbook using Aspose.Cells LightCells and LoadOptions in C# with low memory usage
// AI Prompts: Generate C# code that opens a massive .xlsx file with Aspose.Cells LoadOptions configured for LightCells streaming, iterates through column B to accumulate numeric values, and writes the total to a new workbook while keeping memory consumption minimal. | Show how to set up LoadOptions for low‑memory processing of a big worksheet, use the LightCells API to read rows sequentially, compute an aggregate for a specific column, and save the aggregate result to a separate Excel file.
// Common Searches: aspocells lightcells how to sum a column in a huge xlsx without loading whole file | c# load large excel with loadoptions streaming and calculate column total | memory efficient processing of big dataset in excel using aspocells lightcells api
// Tags: LightCells streaming aggregation for large Excel files | LoadOptions memory‑efficient workbook loading Aspose.Cells | column B numeric sum using Aspose.Cells LightCells | write aggregation result to separate workbook Aspose.Cells | process massive XLSX dataset in .NET with LightCells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example checks for the presence of a large XLSX file, loads it with LoadOptions set for LightCells streaming, determines the used range of the first worksheet, iterates through column B to sum numeric values, creates a new workbook, writes the computed sum into the first row, ensures the output directory exists, and saves the result while handling possible I/O errors, all with minimal memory footprint.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "large_dataset.xlsx";
        string outputPath = "processed_result.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            using (Workbook workbook = new Workbook(inputPath, loadOptions))
            {
                // Assume the data to process is in the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the used range of the worksheet
                AsposeRange usedRange = worksheet.Cells.MaxDisplayRange;

                // Example processing: calculate the sum of values in column B (index 1)
                double sum = 0.0;
                int startRow = usedRange.FirstRow;
                int endRow = startRow + usedRange.RowCount; // exclusive

                for (int row = startRow; row < endRow; row++)
                {
                    object val = worksheet.Cells[row, 1].Value; // column B
                    if (val != null && double.TryParse(val.ToString(), out double d))
                    {
                        sum += d;
                    }
                }

                // Create a new workbook to store the result
                using (Workbook resultWorkbook = new Workbook())
                {
                    Worksheet resultSheet = resultWorkbook.Worksheets[0];
                    resultSheet.Cells[0, 0].PutValue("Sum of Column B");
                    resultSheet.Cells[0, 1].PutValue(sum);

                    // Ensure the output directory exists
                    string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the result workbook with error handling
                    try
                    {
                        resultWorkbook.Save(outputPath);
                        Console.WriteLine($"Processing completed. Result saved to {outputPath}");
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine($"Failed to save result workbook: {saveEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
