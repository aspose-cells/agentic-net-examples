// Title: C# – Replace Substring in All Text Cells of a CSV and Save Cleaned File with Aspose.Cells
// Description: Loads a CSV into a workbook, replaces every occurrence of a given substring in text cells, and exports the result as a new CSV using Aspose.Cells for .NET.
// Keywords: Aspose.Cells CSV replace | C# replace substring CSV | Import CSV Aspose.Cells | Export cleaned CSV C# | Workbook.Replace example | CSV data cleansing Aspose
// Common Searches: replace text in all CSV cells using Aspose.Cells | C# import CSV, replace string, save file | Aspose.Cells replace substring across workbook | how to clean CSV data with Aspose.Cells .NET
// Developer Intent: Load a CSV, substitute a specific substring in every text cell, and write the cleaned data back to a new CSV file using Aspose.Cells for .NET.
// Use Cases: Update legacy codes or terminology across large CSV exports. | Standardize values before importing the file into downstream systems. | Perform bulk data sanitization for reporting or analytics pipelines.
// AI Prompts: Generate C# code that uses Aspose.Cells to replace multiple substrings in a CSV and save the result. | Show how to do a case‑insensitive replacement of a substring in all text cells of a CSV with Aspose.Cells. | Explain how to keep numeric values unchanged while performing text replacement in a CSV using Aspose.Cells.

using System;
using Aspose.Cells;

namespace CsvCleanExample
{
    // Loads a CSV into a workbook, replaces every occurrence of a given substring in text cells, and exports the result as a new CSV using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Paths for input and output CSV files
            string inputCsvPath = "input.csv";
            string outputCsvPath = "cleaned_output.csv";

            // Substring to find and its replacement
            string oldSubstring = "OldValue";
            string newSubstring = "NewValue";

            // Create a new workbook (lifecycle rule: use provided creation)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the CSV file into the worksheet.
            // Using comma as delimiter, keep data as strings (convertNumericData = false)
            cells.ImportCSV(inputCsvPath, ",", false, 0, 0);

            // Replace all occurrences of the specified substring in text cells
            // (Workbook.Replace method handles the replacement across the whole workbook)
            workbook.Replace(oldSubstring, newSubstring);

            // Save the cleaned data back to a CSV file (lifecycle rule: use provided saving)
            workbook.Save(outputCsvPath, SaveFormat.Csv);
        }
    }
}
