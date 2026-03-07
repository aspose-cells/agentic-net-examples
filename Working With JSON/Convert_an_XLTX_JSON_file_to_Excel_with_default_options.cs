using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an XLTX template
            string sourcePath = "template.json";

            // Desired output Excel file (XLSX) path
            string outputPath = "converted.xlsx";

            // Load options specifying that the source file is JSON
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);

            // Save options for the default OOXML Excel format (XLSX)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Perform the conversion using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, loadOptions, outputPath, saveOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{outputPath}'");
        }
    }
}