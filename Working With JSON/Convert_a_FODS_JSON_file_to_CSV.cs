using System;
using Aspose.Cells.Utility;

namespace AsposeCellsFodsToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source FODS file (OpenDocument Flat XML Spreadsheet)
            string sourcePath = "input.fods";

            // Desired output CSV file path
            string destinationPath = "output.csv";

            // Convert the FODS file to CSV using Aspose.Cells ConversionUtility
            // This method handles loading the source format and saving to the target format.
            ConversionUtility.Convert(sourcePath, destinationPath);

            Console.WriteLine($"Conversion completed: \"{sourcePath}\" -> \"{destinationPath}\"");
        }
    }
}