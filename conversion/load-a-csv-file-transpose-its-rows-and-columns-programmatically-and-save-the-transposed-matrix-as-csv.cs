// Title: Programmatically transpose a CSV file’s rows and columns using Aspose.Cells for .NET and export the result as a new CSV
// AI Prompts: Read a CSV file into an Aspose.Cells Workbook, create a range covering the entire data area, call Range.Transpose, and save the transformed worksheet as a CSV using C#. | Write C# code that loads a CSV, flips its matrix with Aspose.Cells’ transpose functionality, and writes the flipped content to another CSV file.
// Common Searches: how to transpose entire CSV file using Aspose.Cells in C# | Aspose.Cells C# example for flipping rows to columns and saving as CSV | convert CSV to workbook, transpose data range, and export back to CSV with Aspose.Cells | C# code to reverse rows and columns of a CSV using Aspose.Cells Range.Transpose
// Tags: Aspose.Cells CSV matrix flip C# | Range.Transpose API Aspose.Cells | export transposed worksheet to CSV .NET | full worksheet range creation Aspose.Cells | CSV to Aspose.Cells workbook conversion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTransposeCsv
{
    // The example loads an input CSV into an Aspose.Cells Workbook, determines the used range, creates a range that covers all populated cells, transposes that range so rows become columns, and saves the resulting matrix to a new CSV file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for input and output CSV files
                string inputCsvPath = "input.csv";
                string outputCsvPath = "output.csv";

                // Verify input file exists
                if (!File.Exists(inputCsvPath))
                {
                    Console.WriteLine($"Input file not found: {inputCsvPath}");
                    return;
                }

                // Load the CSV file into a workbook
                Workbook workbook = new Workbook(inputCsvPath);

                // Access the first worksheet (the CSV data is loaded here)
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Determine the used range dimensions
                int lastRow = cells.MaxDataRow;
                int lastColumn = cells.MaxDataColumn;

                // Create a range that covers all populated cells
                Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

                // Transpose the range (rows become columns and vice versa)
                dataRange.Transpose();

                // Save the transposed data back to a CSV file
                workbook.Save(outputCsvPath, SaveFormat.Csv);
                Console.WriteLine($"Transposed CSV saved to {outputCsvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
