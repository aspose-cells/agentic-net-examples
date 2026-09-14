// Title: Import CSV while skipping lines that start with a custom comment character and save the result as TSV using Aspose.Cells for .NET
// AI Prompts: Read a CSV file, filter out rows whose first non‑blank character matches a given comment symbol, load the remaining data into an Aspose.Cells workbook, and export it as a tab‑separated file. | Use TxtLoadOptions to set a comma delimiter, remove comment lines from a CSV stream, then apply TxtSaveOptions with a tab separator to write the cleaned data to a TSV file in C#.
// Common Searches: c# Aspose.Cells ignore lines starting with # when loading CSV | how to filter comment rows from CSV before converting to TSV with Aspose.Cells | load CSV with custom comment character using TxtLoadOptions Aspose.Cells .NET | convert CSV to TSV after removing comment lines in C# Aspose.Cells | skip commented lines in CSV import Aspose.Cells workbook
// Tags: skip comment lines CSV Aspose.Cells | TxtLoadOptions custom comment character CSV | TxtSaveOptions TSV export Aspose.Cells | remove # lines before CSV import C# | CSV to TSV conversion Aspose.Cells .NET

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsCsvToTsv
{
    // The example reads a CSV file, discards any lines that begin with a specified comment character, loads the filtered data into an Aspose.Cells workbook using TxtLoadOptions, and then saves the workbook as a UTF‑8 encoded TSV file via TxtSaveOptions.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Define the comment character (lines starting with this will be ignored)
            char commentChar = '#';

            // Read all lines, filter out commented ones, and rebuild the content
            string[] allLines = File.ReadAllLines(csvPath);
            string filteredContent = string.Join(Environment.NewLine,
                allLines.Where(line => !line.TrimStart().StartsWith(commentChar.ToString())));

            // Convert the filtered CSV content to a memory stream
            byte[] csvBytes = System.Text.Encoding.UTF8.GetBytes(filteredContent);
            using (MemoryStream csvStream = new MemoryStream(csvBytes))
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Configure load options for CSV (comma as separator)
                TxtLoadOptions loadOptions = new TxtLoadOptions();
                loadOptions.Separator = ',';               // CSV delimiter
                loadOptions.HasTextQualifier = true;       // default behavior
                loadOptions.TreatConsecutiveDelimitersAsOne = false;

                // Import the filtered CSV data starting at cell A1 (row 0, column 0)
                cells.ImportCSV(csvStream, loadOptions, 0, 0);

                // Prepare save options for TSV (tab as separator)
                TxtSaveOptions saveOptions = new TxtSaveOptions();
                saveOptions.Separator = '\t';               // TSV delimiter
                saveOptions.Encoding = System.Text.Encoding.UTF8;

                // Save the workbook as a TSV file
                string tsvPath = "output.tsv";
                workbook.Save(tsvPath, saveOptions);
            }

            Console.WriteLine("CSV data imported (comments ignored) and saved as TSV successfully.");
        }
    }
}
