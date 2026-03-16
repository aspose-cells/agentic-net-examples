using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertXlsxToMultipleTextFormats
    {
        public static void Run()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Convert to CSV
            workbook.Save("output.csv", SaveFormat.Csv);

            // Convert to TSV (tab‑separated values)
            workbook.Save("output.tsv", SaveFormat.Tsv);

            // Convert to TXT using TxtSaveOptions.
            TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
            txtOptions.Separator = '\t'; // Use tab as the delimiter for the TXT file
            workbook.Save("output.txt", txtOptions);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertXlsxToMultipleTextFormats.Run();
        }
    }
}