using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvToTsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Define the custom comment character (lines starting with this will be ignored)
            char commentChar = '#';

            // Read the CSV file, filter out commented lines, and store the clean content in a memory stream
            MemoryStream cleanCsvStream = GetCleanCsvStream(csvPath, commentChar);

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure load options for CSV (comma separator, default text qualifier)
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.Separator = ','; // assuming CSV uses comma as delimiter

            // Import the filtered CSV data into the worksheet (lifecycle rule: load)
            cells.ImportCSV(cleanCsvStream, loadOptions, 0, 0);

            // Prepare save options for TSV (tab separator)
            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.Separator = '\t';

            // Save the workbook as TSV (lifecycle rule: save)
            workbook.Save("output.tsv", saveOptions);

            // Clean up
            cleanCsvStream.Dispose();
        }

        /// <summary>
        /// Reads a CSV file, removes lines that start with the specified comment character,
        /// and returns a MemoryStream containing the cleaned CSV data.
        /// </summary>
        private static MemoryStream GetCleanCsvStream(string filePath, char commentChar)
        {
            // Read all lines from the file
            string[] allLines = File.ReadAllLines(filePath, Encoding.UTF8);

            // Filter out lines that start with the comment character (ignoring leading whitespace)
            StringBuilder sb = new StringBuilder();
            foreach (string line in allLines)
            {
                string trimmed = line.TrimStart();
                if (trimmed.Length == 0 || trimmed[0] != commentChar)
                {
                    sb.AppendLine(line);
                }
            }

            // Convert the cleaned content to a byte array and wrap it in a MemoryStream
            byte[] bytes = Encoding.UTF8.GetBytes(sb.ToString());
            return new MemoryStream(bytes);
        }
    }
}