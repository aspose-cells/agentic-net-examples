// Title: C# – Save Workbook as CSV with UTF‑32 Encoding using Aspose.Cells TxtSaveOptions
// Description: Creates a workbook, fills sample data, sets TxtSaveOptions.Encoding to UTF‑32 and Separator to ',', then saves the file as a UTF‑32 encoded CSV (LargeDataSet.csv).
// Keywords: Aspose.Cells CSV UTF-32 C# | TxtSaveOptions encoding | export workbook to CSV Aspose | UTF-32 CSV large dataset | C# save Excel as UTF-32 CSV
// Common Searches: Aspose.Cells save CSV with UTF-32 C# | TxtSaveOptions UTF-32 encoding example | C# export Excel to CSV using custom encoding | How to set CSV separator in Aspose.Cells | UTF-32 CSV output for large data sets
// Developer Intent: Export an Aspose.Cells workbook to a CSV file encoded in UTF‑32.
// Use Cases: Produce CSV files for systems that require UTF‑32 to retain full Unicode characters. | Generate locale‑independent CSV exports for multilingual datasets. | Create separate UTF‑32 encoded CSV files per worksheet when integrating with legacy pipelines.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a CSV with UTF‑32 encoding and a semicolon delimiter. | Explain when UTF‑32 should be preferred over UTF‑8 for CSV exports and its impact on file size. | Provide a step‑by‑step guide to export each worksheet of a workbook to individual UTF‑32 encoded CSV files using Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvUtf32Example
{
    // Creates a workbook, fills sample data, sets TxtSaveOptions.Encoding to UTF‑32 and Separator to ',', then saves the file as a UTF‑32 encoded CSV (LargeDataSet.csv).
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Create CSV (text) save options
            TxtSaveOptions saveOptions = new TxtSaveOptions();

            // Specify UTF‑32 encoding for large datasets
            saveOptions.Encoding = Encoding.UTF32;

            // Optionally set the separator to comma (default for CSV)
            saveOptions.Separator = ',';

            // Save the workbook as CSV using the custom options
            workbook.Save("LargeDataSet.csv", saveOptions);

            Console.WriteLine("Workbook saved as CSV with UTF‑32 encoding.");
        }
    }
}
