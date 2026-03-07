using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Utility;      // For JsonSaveOptions

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLTX template file
            string sourcePath = "template.xltx";

            // Path where the resulting JSON will be saved
            string jsonPath = "output.json";

            // Load the XLTX workbook (default load options are sufficient)
            Workbook workbook = new Workbook(sourcePath);

            // Create default JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON using the default options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Conversion completed. JSON saved to: {jsonPath}");
        }
    }
}