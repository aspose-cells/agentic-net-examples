// Title: C# – Load CSV with custom comment character, filter comments, and export to TSV using Aspose.Cells
// Description: A concise example that reads a CSV file, removes empty rows and any line whose first non‑space character matches a user‑defined comment marker (e.g., '#'), imports the cleaned data into an Aspose.Cells workbook with TxtLoadOptions, and saves the worksheet as a UTF‑8 TSV file using TxtSaveOptions.
// Keywords: Aspose.Cells CSV import C# | custom comment character CSV | skip commented lines Aspose | CSV to TSV conversion .NET | TxtLoadOptions comment filter | TxtSaveOptions TSV export | C# read CSV ignore # lines | Aspose.Cells data cleaning
// Common Searches: How to ignore comment lines when loading CSV with Aspose.Cells | Convert CSV to TSV in C# while skipping # rows | Aspose.Cells load CSV with custom comment character | C# filter commented rows before TSV export | Aspose.Cells TxtLoadOptions comment handling
// Developer Intent: Read a CSV, discard lines that start with a specified comment character, and write the remaining data to a TSV file using Aspose.Cells.
// Use Cases: Pre‑process configuration or log files that contain comment rows before feeding them to analytics pipelines. | Clean exported data sets by removing header comments prior to bulk import into databases. | Automate transformation of comment‑annotated CSV reports into tab‑delimited format for spreadsheet or BI tools.
// AI Prompts: Generate C# code with Aspose.Cells to load a CSV, skip lines beginning with '#', and save the result as a TSV file. | Explain how TxtLoadOptions and TxtSaveOptions can be set to handle custom comment characters and delimiters in Aspose.Cells. | Suggest memory‑efficient techniques for filtering large CSV files with comment lines before importing them with Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvToTsv
{
    // A concise example that reads a CSV file, removes empty rows and any line whose first non‑space character matches a user‑defined comment marker (e.g., '#'), imports the cleaned data into an Aspose.Cells workbook with TxtLoadOptions, and saves the worksheet as a UTF‑8 TSV file using TxtSaveOptions.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Define the custom comment character (e.g., '#')
            char commentChar = '#';

            // Read the CSV file, filter out commented lines, and build a clean CSV string
            string[] allLines = File.ReadAllLines(csvPath);
            StringBuilder cleanBuilder = new StringBuilder();
            foreach (string line in allLines)
            {
                // Trim leading spaces before checking the comment character
                string trimmed = line.TrimStart();
                if (trimmed.Length == 0) continue;               // Skip empty lines
                if (trimmed[0] == commentChar) continue;         // Skip comment lines
                cleanBuilder.AppendLine(line);
            }

            // Convert the cleaned CSV content to a memory stream
            byte[] cleanBytes = Encoding.UTF8.GetBytes(cleanBuilder.ToString());
            using (MemoryStream cleanStream = new MemoryStream(cleanBytes))
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Configure load options for CSV (comma separator, convert numeric data)
                TxtLoadOptions loadOptions = new TxtLoadOptions();
                loadOptions.Separator = ',';          // CSV delimiter
                loadOptions.ConvertNumericData = true;

                // Import the cleaned CSV data starting at cell A1 (row 0, column 0)
                cells.ImportCSV(cleanStream, loadOptions, 0, 0);

                // Prepare save options for TSV (tab delimiter)
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv);
                saveOptions.Separator = '\t';         // TSV delimiter
                saveOptions.Encoding = Encoding.UTF8;

                // Save the workbook as a TSV file
                workbook.Save("output.tsv", saveOptions);
            }

            Console.WriteLine("CSV imported (comments ignored) and saved as TSV successfully.");
        }
    }
}
