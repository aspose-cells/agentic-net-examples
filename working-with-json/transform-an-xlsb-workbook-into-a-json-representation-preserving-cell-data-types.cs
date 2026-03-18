using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSB workbook
            string sourcePath = "input.xlsb";

            // Load the XLSB workbook (Aspose.Cells detects format from extension)
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options to preserve Excel structure and data types
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON representation of the Excel structure
                ToExcelStruct = true,

                // Include empty cells in the output (null values)
                ExportEmptyCells = true,

                // Keep original data types (do not force all values to strings)
                ExportAsString = false
            };

            // Save the workbook as a JSON file using the configured options
            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);
        }
    }
}