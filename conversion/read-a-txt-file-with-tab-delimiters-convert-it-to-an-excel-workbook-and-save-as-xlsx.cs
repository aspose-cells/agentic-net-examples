using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace TxtToExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source tab‑delimited text file
            string sourcePath = "data.txt";

            // Desired output Excel file path
            string outputPath = "data.xlsx";

            // Load options specifying that the source file is a TSV (tab‑separated) file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);

            // Save options for the XLSX format
            SaveOptions saveOptions = new OoxmlSaveOptions();

            // Perform the conversion using the provided ConversionUtility rule
            ConversionUtility.Convert(sourcePath, loadOptions, outputPath, saveOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' → '{outputPath}'");
        }
    }
}