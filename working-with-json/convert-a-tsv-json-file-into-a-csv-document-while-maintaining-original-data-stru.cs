using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    class TsvJsonToCsv
    {
        static void Main()
        {
            // Path to the source TSV (tab‑separated) file that contains JSON data
            string tsvFilePath = "input.tsv";

            // Path for the resulting CSV file
            string csvFilePath = "output.csv";

            // -----------------------------------------------------------------
            // Load the TSV file
            // -----------------------------------------------------------------
            // Create load options for a tab‑delimited text file.
            // Set the separator to a tab character and preserve UTF‑8 encoding.
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
            loadOptions.Separator = '\t';
            loadOptions.Encoding = Encoding.UTF8;   // keep original encoding

            // Load the TSV content into a workbook.
            Workbook workbook = new Workbook(tsvFilePath, loadOptions);

            // -----------------------------------------------------------------
            // Save as CSV
            // -----------------------------------------------------------------
            // Create save options for a comma‑separated CSV file.
            // Keep the same UTF‑8 encoding to maintain data integrity.
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
            saveOptions.Separator = ',';           // CSV delimiter
            saveOptions.Encoding = Encoding.UTF8;  // preserve encoding

            // Save the workbook as CSV.
            workbook.Save(csvFilePath, saveOptions);

            Console.WriteLine($"TSV file '{tsvFilePath}' has been converted to CSV '{csvFilePath}'.");
        }
    }
}