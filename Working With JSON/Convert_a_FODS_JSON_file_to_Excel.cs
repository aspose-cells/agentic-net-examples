using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source FODS file (OpenDocument Flat XML Spreadsheet)
            string sourcePath = "input.fods";

            // Desired output Excel file path (XLSX)
            string destinationPath = "output.xlsx";

            // LoadOptions specifying that the source format is FODS
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);

            // SaveOptions for saving as the default Office Open XML format (XLSX)
            SaveOptions saveOptions = new OoxmlSaveOptions();

            // Perform the conversion using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, loadOptions, destinationPath, saveOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }
}