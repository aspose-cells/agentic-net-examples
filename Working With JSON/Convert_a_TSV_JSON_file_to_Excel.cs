using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source TSV file (JSON formatted TSV data)
            string sourcePath = "input.tsv";

            // Desired output Excel file path
            string outputPath = "output.xlsx";

            // Load options specifying that the source file is a TSV (Tab‑Separated Values) file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);

            // Save options for creating a standard XLSX workbook
            SaveOptions saveOptions = new OoxmlSaveOptions();

            // Convert the TSV file to Excel using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, loadOptions, outputPath, saveOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' → '{outputPath}'");
        }
    }
}