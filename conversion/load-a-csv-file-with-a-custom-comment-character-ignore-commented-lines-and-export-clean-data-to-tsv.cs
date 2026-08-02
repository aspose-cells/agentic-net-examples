// Title: C# – Load CSV, skip custom comment lines, and export clean data to TSV with Aspose.Cells
// Description: Shows how to read a CSV file, filter out rows that start with a user‑defined comment character (e.g., '#'), import the filtered content into an Aspose.Cells workbook via a MemoryStream using TxtLoadOptions, and save it as a tab‑delimited TSV file with TxtSaveOptions.
// Keywords: Aspose.Cells | C# CSV to TSV | ignore comment lines | custom comment character | TxtLoadOptions | TxtSaveOptions | MemoryStream CSV import | CSV preprocessing | tab delimited export
// Common Searches: Aspose.Cells ignore commented rows when loading CSV | C# convert CSV to TSV with Aspose.Cells | filter # comments from CSV before saving as TSV | load CSV from MemoryStream Aspose.Cells .NET | custom separator CSV Aspose.Cells example
// Developer Intent: Read a CSV, discard lines that begin with a specified comment marker, and write the remaining data to a TSV file using Aspose.Cells for .NET.
// Use Cases: Cleaning log or data files that contain comment lines before analysis. | Automating CSV‑to‑TSV conversion in ETL pipelines where only non‑comment rows are required. | Preparing tab‑delimited files for import into databases or BI tools after removing metadata comments.
// AI Prompts: Generate C# code that uses Aspose.Cells to load a CSV from a MemoryStream, ignore rows starting with a given comment character, and save the result as a TSV file. | Explain how TxtLoadOptions and TxtSaveOptions can be configured to handle custom separators and comment filtering during CSV‑to‑TSV conversion with Aspose.Cells. | Suggest alternative methods (e.g., StreamReader, custom parser) for removing comment lines before importing a CSV into Aspose.Cells.

using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvToTsv
{
    // Shows how to read a CSV file, filter out rows that start with a user‑defined comment character (e.g., '#'), import the filtered content into an Aspose.Cells workbook via a MemoryStream using TxtLoadOptions, and save it as a tab‑delimited TSV file with TxtSaveOptions.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Define the comment character (lines starting with this will be ignored)
            char commentChar = '#';

            // Read all lines, filter out commented lines, and rebuild the CSV content
            string[] allLines = File.ReadAllLines(csvPath);
            string[] dataLines = allLines
                .Where(line => !line.TrimStart().StartsWith(commentChar.ToString()))
                .ToArray();
            string filteredCsv = string.Join(Environment.NewLine, dataLines);

            // Load the filtered CSV data into a memory stream
            using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(filteredCsv)))
            {
                // Configure load options for CSV (comma as separator)
                TxtLoadOptions loadOptions = new TxtLoadOptions();
                loadOptions.Separator = ',';
                loadOptions.ConvertNumericData = true; // optional: convert numbers

                // Create a new workbook and import the CSV data
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells.ImportCSV(csvStream, loadOptions, 0, 0);

                // Prepare save options for TSV (tab as separator)
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv);
                saveOptions.Separator = '\t';

                // Save the cleaned data as TSV
                workbook.Save("output.tsv", saveOptions);
            }
        }
    }
}
