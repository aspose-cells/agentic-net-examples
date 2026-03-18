using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLTX template file
            string sourcePath = "template.xltx";

            // Desired output JSON file path
            string outputPath = "output.json";

            // Convert the XLTX file to JSON using default conversion parameters
            ConversionUtility.Convert(sourcePath, outputPath);

            Console.WriteLine($"Conversion completed. JSON saved to: {outputPath}");
        }
    }
}