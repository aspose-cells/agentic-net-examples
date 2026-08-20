// Title: Transpose a CSV file with Aspose.Cells in C# and save as a new CSV
// AI Prompts: Write C# code that loads a CSV into an Aspose.Cells Workbook, transposes the data range, and writes the result to another CSV file. | Show how to check for the existence of the source CSV and handle empty‑file scenarios before performing a transpose with Aspose.Cells. | Provide example error handling for file‑not‑found and invalid data errors when transposing CSV data using Aspose.Cells.
// Common Searches: Aspose.Cells transpose CSV C# | C# transpose rows and columns of a CSV file | How to save transposed data to CSV with Aspose.Cells | Validate CSV file before processing with Aspose.Cells | Range.Transpose example in Aspose.Cells
// Tags: Aspose.Cells | C# | CSV transpose | Range.Transpose | Error handling

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the input CSV exists, loads it into an Aspose.Cells Workbook using CSV load options, transposes the populated range, and saves the transformed matrix to a new CSV file while handling missing‑file and empty‑data errors.
class CsvTranspose
{
    static void Main()
    {
        // Paths for input and output CSV files
        string inputCsv = "input.csv";
        string outputCsv = "transposed.csv";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputCsv))
                throw new FileNotFoundException($"Input file not found: {inputCsv}");

            // Load the CSV file into a workbook using CSV load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputCsv, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Get the range that contains all populated cells
            Aspose.Cells.Range dataRange = cells.MaxDisplayRange;

            // Ensure there is data to transpose
            if (dataRange == null || dataRange.RowCount == 0 || dataRange.ColumnCount == 0)
                throw new InvalidOperationException("The input CSV does not contain any data to transpose.");

            // Transpose the data (swap rows and columns)
            dataRange.Transpose();

            // Save the transposed data back to a CSV file
            workbook.Save(outputCsv, SaveFormat.Csv);
            Console.WriteLine($"Transposition completed. Output saved to '{outputCsv}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
