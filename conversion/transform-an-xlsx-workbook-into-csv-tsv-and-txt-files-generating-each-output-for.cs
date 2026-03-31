using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "source.xlsx";

            // Output file paths for each format
            string csvPath = "output.csv";
            string tsvPath = "output.tsv";
            string txtPath = "output.txt";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Save as CSV (comma‑separated values)
            workbook.Save(csvPath, SaveFormat.Csv);

            // Save as TSV (tab‑separated values)
            workbook.Save(tsvPath, SaveFormat.Tsv);

            // Save as a generic text file using TxtSaveOptions.
            // Here we use CSV format with a custom separator (tab) to produce a TXT file.
            TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
            txtOptions.SeparatorString = "\t";   // Use tab as the delimiter
            workbook.Save(txtPath, txtOptions);

            Console.WriteLine("Conversion completed: CSV, TSV, and TXT files generated.");
        }
    }
}