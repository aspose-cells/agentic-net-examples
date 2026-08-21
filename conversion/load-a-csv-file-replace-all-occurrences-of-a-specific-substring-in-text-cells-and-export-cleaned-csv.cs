// Title: C# – Import CSV, Replace Substring in All Cells, Export Clean CSV with Aspose.Cells
// Description: Load a CSV into an Aspose.Cells Workbook, replace every occurrence of a given substring in text cells using Workbook.Replace, report the replacement count, and save the cleaned data as a new CSV file.
// Keywords: Aspose.Cells CSV import | Aspose.Cells replace text | C# CSV cleaning | Workbook.Replace .NET | export CSV Aspose | data sanitization C# | text substitution Excel library
// Common Searches: Aspose.Cells replace string in CSV C# | How to clean CSV data with Aspose.Cells | C# import CSV and replace values Aspose | Workbook.Replace example for CSV files | Export modified CSV using Aspose.Cells .NET
// Developer Intent: Replace a target substring in every text cell of a CSV file and write the sanitized result to a new CSV.
// Use Cases: Remove confidential identifiers before sharing CSV exports. | Standardize terminology across legacy data sets. | Prepare CSV files for ETL pipelines by eliminating outdated values.
// AI Prompts: Write C# code that uses Aspose.Cells to load a CSV, replace "oldValue" with "newValue" in all cells, and save the output as a new CSV. | Explain the behavior of Workbook.Replace on numeric versus text cells after importing a CSV with Aspose.Cells. | Create a script that processes multiple CSV files, applying the same substring replacement using Aspose.Cells in a loop.

using System;
using Aspose.Cells;

namespace CsvCleaner
{
    // Load a CSV into an Aspose.Cells Workbook, replace every occurrence of a given substring in text cells using Workbook.Replace, report the replacement count, and save the cleaned data as a new CSV file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input CSV file path
            string inputCsvPath = "input.csv";

            // Output CSV file path (cleaned)
            string outputCsvPath = "cleaned.csv";

            // Substring to find and its replacement
            string oldSubstring = "oldValue";
            string newSubstring = "newValue";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the CSV data into the worksheet starting at cell A1 (row 0, column 0)
            // Using comma as delimiter and converting numeric data where possible
            cells.ImportCSV(inputCsvPath, ",", true, 0, 0);

            // Replace all occurrences of the specified substring in all text cells
            int replacedCount = workbook.Replace(oldSubstring, newSubstring);
            Console.WriteLine($"Total replacements made: {replacedCount}");

            // Save the cleaned data back to a CSV file
            workbook.Save(outputCsvPath, SaveFormat.Csv);
        }
    }
}
