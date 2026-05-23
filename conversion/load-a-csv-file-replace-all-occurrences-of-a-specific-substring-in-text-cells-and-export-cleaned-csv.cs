using System;
using Aspose.Cells;

namespace CsvCleaner
{
    class Program
    {
        static void Main()
        {
            // Paths to the input and output CSV files
            string inputCsvPath = "input.csv";
            string outputCsvPath = "cleaned_output.csv";

            // Substring to find and its replacement
            string oldSubstring = "OldValue";
            string newSubstring = "NewValue";

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Import the CSV data into the first worksheet starting at cell A1
            // Using comma as the delimiter and converting numeric data where possible
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells.ImportCSV(inputCsvPath, ",", true, 0, 0); // lifecycle rule: load

            // Replace all occurrences of the specified substring in all text cells
            workbook.Replace(oldSubstring, newSubstring);

            // Save the cleaned data back to a CSV file (lifecycle rule: save)
            workbook.Save(outputCsvPath, SaveFormat.Csv);
        }
    }
}